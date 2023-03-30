import pandas
from flask_sqlalchemy import SQLAlchemy
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root:@localhost/codingthunder'
db = SQLAlchemy(app)
import pandas
from flask_sqlalchemy import SQLAlchemy
app.config['SQLALCHEMY_DATABASE_URI'] = 'mysql://root:@localhost/codingthunder'
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['UPLOAD_FOLDER'] = r"C:\Users\AMIT\OneDrive\Desktop\Arjun_Sir\Upload"
db = SQLAlchemy(app)

class Forecast(db.Model):
    sno = db.Column(db.Integer, primary_key=True)
    Hotel = db.Column(db.String(80), nullable=False)
    Month = db.Column(db.String(12), nullable=False)
    Quarter = db.Column(db.String(120), nullable=False)
    Occupancy_Date = db.Column(db.String(12), nullable=True)
    DOW = db.Column(db.String(20), nullable=False)
    RIS_FIT = db.Column(db.Integer, nullable=False)
    RS_Group = db.Column(db.Integer, nullable=False)
    RS_Comp_H_Use = db.Column(db.Integer, nullable=False)
    Total_Rooms_Sold = db.Column(db.Integer, nullable=False)
    FC_FIT = db.Column(db.Integer, nullable=False)
    FC_Group = db.Column(db.Integer, nullable=False)
    FC_Comp_H_Use = db.Column(db.Integer, nullable=False)
    Total_Forecast = db.Column(db.Integer, nullable=False)

db = SQLAlchemy(app)
entry = Forecast(name=name, phone_num=phone,
                         msg=message, date=datetime.now(), email=email)
db.session.add(entry)
db.session.commit()