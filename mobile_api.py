"""Additive mobile API for E-2 School ERP.
No existing HTML route is replaced. The API is designed for the future Android/iOS client.
"""
import base64, hashlib, hmac, json, os, time
from datetime import date
from functools import wraps
from flask import request, jsonify
from werkzeug.security import check_password_hash


def _b64e(raw: bytes) -> str:
    return base64.urlsafe_b64encode(raw).decode().rstrip("=")


def _b64d(value: str) -> bytes:
    return base64.urlsafe_b64decode(value + "=" * (-len(value) % 4))


def register_mobile_api(app, db, models):
    Institute=models["Institute"]; User=models["User"]; Student=models["Student"]
    Fee=models["Fee"]; Attendance=models["Attendance"]; Transaction=models["Transaction"]
    ExamSession=models["ExamSession"]

    def secret():
        return (app.config.get("SECRET_KEY") or os.environ.get("E2_SECRET_KEY") or "").encode()

    def make_token(user):
        payload={"uid":user.id,"iid":user.institute_id,"role":user.role,"exp":int(time.time())+60*60*24*30}
        body=_b64e(json.dumps(payload,separators=(",",":"),sort_keys=True).encode())
        sig=_b64e(hmac.new(secret(),body.encode(),hashlib.sha256).digest())
        return body+"."+sig

    def read_token(token):
        try:
            body,sig=token.split(".",1)
            expected=_b64e(hmac.new(secret(),body.encode(),hashlib.sha256).digest())
            if not hmac.compare_digest(sig,expected): return None
            payload=json.loads(_b64d(body))
            if int(payload.get("exp",0)) < int(time.time()): return None
            return payload
        except Exception:
            return None

    def api_auth(fn):
        @wraps(fn)
        def wrapper(*args,**kwargs):
            if request.method=="OPTIONS": return ("",204)
            header=request.headers.get("Authorization","")
            token=header[7:].strip() if header.lower().startswith("bearer ") else ""
            payload=read_token(token)
            if not payload:
                return jsonify({"ok":False,"error":"Unauthorized"}),401
            user=User.query.filter_by(id=int(payload["uid"])).first()
            if not user or not user.institute_id or int(user.institute_id)!=int(payload["iid"]):
                return jsonify({"ok":False,"error":"Unauthorized"}),401
            return fn(payload,user,*args,**kwargs)
        return wrapper

    @app.after_request
    def mobile_api_headers(response):
        if request.path.startswith("/api/"):
            origin=os.environ.get("E2_API_ORIGIN","*")
            response.headers["Access-Control-Allow-Origin"]=origin
            response.headers["Access-Control-Allow-Headers"]="Content-Type, Authorization"
            response.headers["Access-Control-Allow-Methods"]="GET, POST, PUT, PATCH, DELETE, OPTIONS"
            response.headers["Cache-Control"]="no-store"
        return response

    @app.route("/api/v1/health",methods=["GET","OPTIONS"])
    def api_health():
        if request.method=="OPTIONS": return ("",204)
        try:
            db.session.execute(db.text("SELECT 1"))
            return jsonify({"ok":True,"service":"E-2 School ERP API","version":"1.0"})
        except Exception:
            return jsonify({"ok":False,"error":"Database unavailable"}),503

    @app.route("/api/v1/auth/login",methods=["POST","OPTIONS"])
    def api_login():
        if request.method=="OPTIONS": return ("",204)
        data=request.get_json(silent=True) or {}
        email=str(data.get("email","")).strip().lower()
        password=str(data.get("password", ""))
        if not email or not password:
            return jsonify({"ok":False,"error":"Email and password are required"}),400
        user=User.query.filter(db.func.lower(User.email)==email).first()
        if not user or not check_password_hash(user.password_hash,password):
            return jsonify({"ok":False,"error":"Invalid login"}),401

        # Super-admin accounts created by the ERP can legitimately have no
        # institute_id. For mobile, bind the session to the first active
        # institute so the existing dashboard/data isolation remains intact.
        if not user.institute_id and user.role == "superadmin":
            inst=Institute.query.filter_by(active=True).order_by(Institute.id).first()
            if not inst:
                return jsonify({"ok":False,"error":"Create/activate an institute before using the mobile dashboard"}),409
            mobile_iid=inst.id
        else:
            mobile_iid=user.institute_id
            inst=Institute.query.filter_by(id=mobile_iid).first() if mobile_iid else None
            if not mobile_iid or not inst:
                return jsonify({"ok":False,"error":"User is not assigned to an active institute"}),403

        # Keep the original user unchanged; token carries the mobile institute
        # context for this session.
        original_iid=user.institute_id
        if original_iid == mobile_iid:
            token=make_token(user)
        else:
            payload={"uid":user.id,"iid":mobile_iid,"role":user.role,"exp":int(time.time())+60*60*24*30}
            body=_b64e(json.dumps(payload,separators=(",",":"),sort_keys=True).encode())
            sig=_b64e(hmac.new(secret(),body.encode(),hashlib.sha256).digest())
            token=body+"."+sig

        return jsonify({"ok":True,"token":token,"user":{"id":user.id,"email":user.email,"role":user.role},"institute":{"id":inst.id,"name":inst.name}})

    @app.route("/api/v1/me",methods=["GET","OPTIONS"])
    @api_auth
    def api_me(payload,user):
        inst=Institute.query.filter_by(id=user.institute_id).first()
        return jsonify({"ok":True,"user":{"id":user.id,"email":user.email,"role":user.role},"institute":{"id":inst.id,"name":inst.name} if inst else None})

    @app.route("/api/v1/dashboard",methods=["GET","OPTIONS"])
    @api_auth
    def api_dashboard(payload,user):
        iid=user.institute_id
        students=Student.query.filter_by(institute_id=iid).all()
        fees=Fee.query.filter_by(institute_id=iid).all()
        attendance=Attendance.query.filter_by(institute_id=iid).all()
        tx=Transaction.query.filter_by(institute_id=iid).all()
        today=date.today()
        today_att=[a for a in attendance if a.attendance_date==today]
        today_present=sum(1 for a in today_att if (a.status or '').lower()=='present')
        pending=sum(max(0,float(f.amount or 0)-float(f.discount or 0)-float(f.paid or 0)) for f in fees)
        income=sum(float(t.amount or 0) for t in tx if (t.kind or '').lower()=='income')
        expenses=sum(float(t.amount or 0) for t in tx if (t.kind or '').lower()=='expense')
        return jsonify({"ok":True,"data":{"students_total":len(students),"students_active":sum(1 for s in students if (s.status or '').lower()=='active'),"fees_received":sum(float(f.paid or 0) for f in fees),"fees_pending":pending,"today_attendance_total":len(today_att),"today_attendance_present":today_present,"income_total":income,"expenses_total":expenses,"balance":income-expenses,"exam_sessions":ExamSession.query.filter_by(institute_id=iid).count()}})

    @app.route("/api/v1/students",methods=["GET","OPTIONS"])
    @api_auth
    def api_students(payload,user):
        iid=user.institute_id
        q=(request.args.get("q") or "").strip()
        query=Student.query.filter_by(institute_id=iid).order_by(Student.name)
        if q:
            like=f"%{q}%"
            query=query.filter(db.or_(Student.name.ilike(like),Student.admission_no.ilike(like),Student.father_name.ilike(like),Student.phone.ilike(like)))
        page=max(1,int(request.args.get("page",1) or 1)); per=min(100,max(1,int(request.args.get("per_page",50) or 50)))
        total=query.count(); rows=query.offset((page-1)*per).limit(per).all()
        return jsonify({"ok":True,"page":page,"per_page":per,"total":total,"items":[{"id":s.id,"admission_no":s.admission_no,"name":s.name,"father_name":s.father_name,"phone":s.phone,"class_name":s.class_name,"section":s.section,"monthly_fee":s.monthly_fee,"status":s.status} for s in rows]})

    @app.route("/api/v1/students/<int:sid>",methods=["GET","OPTIONS"])
    @api_auth
    def api_student(payload,user,sid):
        s=Student.query.filter_by(id=sid,institute_id=user.institute_id).first()
        if not s: return jsonify({"ok":False,"error":"Student not found"}),404
        fees=Fee.query.filter_by(student_id=s.id,institute_id=user.institute_id).order_by(Fee.id.desc()).all()
        return jsonify({"ok":True,"item":{"id":s.id,"admission_no":s.admission_no,"name":s.name,"father_name":s.father_name,"phone":s.phone,"class_name":s.class_name,"section":s.section,"monthly_fee":s.monthly_fee,"status":s.status,"fees":[{"id":f.id,"month":f.month,"amount":f.amount,"discount":f.discount,"paid":f.paid,"paid_on":f.paid_on.isoformat() if f.paid_on else None,"note":f.note} for f in fees]}})
