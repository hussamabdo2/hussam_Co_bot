# ============================================================
# MEDICAL REPRESENTATIVE MANAGEMENT BOT
# Python + Telegram + PostgreSQL
# نسخة أولية متكاملة وقابلة للتوسع
# ============================================================

import os
import logging
from datetime import datetime, date, timedelta
from enum import Enum

from dotenv import load_dotenv

from sqlalchemy import (
    create_engine,
    Column,
    Integer,
    BigInteger,
    String,
    Text,
    Date,
    DateTime,
    Float,
    Boolean,
    ForeignKey,
    UniqueConstraint,
    func,
)
from sqlalchemy.orm import (
    declarative_base,
    relationship,
    sessionmaker,
)

from telegram import (
    Update,
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    ReplyKeyboardRemove,
)

from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ConversationHandler,
    ContextTypes,
    filters,
)

# ============================================================
# الإعدادات
# ============================================================

load_dotenv()

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
DATABASE_URL = os.getenv(
    "DATABASE_URL",
    "postgresql://postgres:postgres@localhost:5432/medical_rep_bot"
)

COMPANY_NAME = os.getenv("COMPANY_NAME", "الشركة الدوائية")

ADMIN_IDS_TEXT = os.getenv("ADMIN_IDS", "")

ADMIN_IDS = []

for value in ADMIN_IDS_TEXT.split(","):
    value = value.strip()

    if value.isdigit():
        ADMIN_IDS.append(int(value))


if not BOT_TOKEN:
    raise ValueError(
        "BOT_TOKEN غير موجود. "
        "قم بإضافته في متغيرات البيئة."
    )


# ============================================================
# السجل
# ============================================================

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO
)

logger = logging.getLogger(__name__)


# ============================================================
# قاعدة البيانات
# ============================================================

Base = declarative_base()

import os
from sqlalchemy import create_engine

DATABASE_URL = os.getenv("DATABASE_URL")

if not DATABASE_URL:
    raise RuntimeError(
        "DATABASE_URL غير موجود في Railway."
    )

if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace(
        "postgres://",
        "postgresql://",
        1
    )

engine = create_engine(
    DATABASE_URL,
    pool_pre_ping=True,
    future=True
)


# ============================================================
# الثوابت
# ============================================================

ROLE_ADMIN = "admin"
ROLE_MANAGER = "manager"
ROLE_REPRESENTATIVE = "representative"
ROLE_WAREHOUSE = "warehouse"
ROLE_PHARMACY_OWNER = "pharmacy_owner"
ROLE_PHARMACIST = "pharmacist"


DOCTOR_CATEGORY_A = "A"
DOCTOR_CATEGORY_B = "B"
DOCTOR_CATEGORY_C = "C"
DOCTOR_CATEGORY_D = "D"


# ============================================================
# جداول المستخدمين
# ============================================================

class User(Base):

    __tablename__ = "users"

    id = Column(
        Integer,
        primary_key=True
    )

    telegram_id = Column(
        BigInteger,
        unique=True,
        nullable=False,
        index=True
    )

    full_name = Column(
        String(255),
        nullable=True
    )

    username = Column(
        String(255),
        nullable=True
    )

    phone = Column(
        String(50),
        nullable=True
    )

    role = Column(
        String(50),
        default=ROLE_REPRESENTATIVE
    )

    is_active = Column(
        Boolean,
        default=True
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )

    updated_at = Column(
        DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow
    )


# ============================================================
# التخصصات
# ============================================================

class Specialty(Base):

    __tablename__ = "specialties"

    id = Column(
        Integer,
        primary_key=True
    )

    name = Column(
        String(255),
        unique=True,
        nullable=False
    )

    description = Column(
        Text,
        nullable=True
    )


# ============================================================
# المستشفيات
# ============================================================

class Hospital(Base):

    __tablename__ = "hospitals"

    id = Column(
        Integer,
        primary_key=True
    )

    name = Column(
        String(255),
        nullable=False,
        index=True
    )

    hospital_type = Column(
        String(100),
        nullable=True
    )

    governorate = Column(
        String(100),
        nullable=True
    )

    city = Column(
        String(100),
        nullable=True
    )

    district = Column(
        String(100),
        nullable=True
    )

    address = Column(
        Text,
        nullable=True
    )

    phone = Column(
        String(50),
        nullable=True
    )

    notes = Column(
        Text,
        nullable=True
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )


# ============================================================
# العيادات
# ============================================================

class Clinic(Base):

    __tablename__ = "clinics"

    id = Column(
        Integer,
        primary_key=True
    )

    name = Column(
        String(255),
        nullable=False
    )

    governorate = Column(
        String(100),
        nullable=True
    )

    city = Column(
        String(100),
        nullable=True
    )

    district = Column(
        String(100),
        nullable=True
    )

    address = Column(
        Text,
        nullable=True
    )

    phone = Column(
        String(50),
        nullable=True
    )


# ============================================================
# الأطباء
# ============================================================

class Doctor(Base):

    __tablename__ = "doctors"

    id = Column(
        Integer,
        primary_key=True
    )

    full_name = Column(
        String(255),
        nullable=False,
        index=True
    )

    phone = Column(
        String(50),
        nullable=True,
        index=True
    )

    specialty_id = Column(
        Integer,
        ForeignKey("specialties.id"),
        nullable=True
    )

    category = Column(
        String(10),
        default=DOCTOR_CATEGORY_D
    )

    scientific_degree = Column(
        String(255),
        nullable=True
    )

    governorate = Column(
        String(100),
        nullable=True
    )

    city = Column(
        String(100),
        nullable=True
    )

    district = Column(
        String(100),
        nullable=True
    )

    address = Column(
        Text,
        nullable=True
    )

    working_days = Column(
        String(255),
        nullable=True
    )

    working_hours = Column(
        String(255),
        nullable=True
    )

    importance_score = Column(
        Float,
        default=0
    )

    prescription_score = Column(
        Float,
        default=0
    )

    last_visit_date = Column(
        Date,
        nullable=True
    )

    next_visit_date = Column(
        Date,
        nullable=True
    )

    assigned_rep_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=True
    )

    notes = Column(
        Text,
        nullable=True
    )

    is_active = Column(
        Boolean,
        default=True
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )

    specialty = relationship("Specialty")

    assigned_rep = relationship(
        "User",
        foreign_keys=[assigned_rep_id]
    )


# ============================================================
# ربط الطبيب بالمستشفى
# ============================================================

class DoctorHospital(Base):

    __tablename__ = "doctor_hospitals"

    id = Column(
        Integer,
        primary_key=True
    )

    doctor_id = Column(
        Integer,
        ForeignKey("doctors.id"),
        nullable=False
    )

    hospital_id = Column(
        Integer,
        ForeignKey("hospitals.id"),
        nullable=False
    )

    department = Column(
        String(255),
        nullable=True
    )

    doctor = relationship("Doctor")

    hospital = relationship("Hospital")

    __table_args__ = (
        UniqueConstraint(
            "doctor_id",
            "hospital_id",
            name="uq_doctor_hospital"
        ),
    )


# ============================================================
# ربط الطبيب بالعيادة
# ============================================================

class DoctorClinic(Base):

    __tablename__ = "doctor_clinics"

    id = Column(
        Integer,
        primary_key=True
    )

    doctor_id = Column(
        Integer,
        ForeignKey("doctors.id"),
        nullable=False
    )

    clinic_id = Column(
        Integer,
        ForeignKey("clinics.id"),
        nullable=False
    )

    doctor = relationship("Doctor")

    clinic = relationship("Clinic")


# ============================================================
# الصيدليات
# ============================================================

class Pharmacy(Base):

    __tablename__ = "pharmacies"

    id = Column(
        Integer,
        primary_key=True
    )

    name = Column(
        String(255),
        nullable=False,
        index=True
    )

    owner_name = Column(
        String(255),
        nullable=True
    )

    phone = Column(
        String(50),
        nullable=True
    )

    governorate = Column(
        String(100),
        nullable=True
    )

    city = Column(
        String(100),
        nullable=True
    )

    district = Column(
        String(100),
        nullable=True
    )

    address = Column(
        Text,
        nullable=True
    )

    latitude = Column(
        Float,
        nullable=True
    )

    longitude = Column(
        Float,
        nullable=True
    )

    classification = Column(
        String(50),
        default="عادية"
    )

    assigned_rep_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=True
    )

    last_visit_date = Column(
        Date,
        nullable=True
    )

    next_visit_date = Column(
        Date,
        nullable=True
    )

    notes = Column(
        Text,
        nullable=True
    )

    is_active = Column(
        Boolean,
        default=True
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )

    assigned_rep = relationship(
        "User",
        foreign_keys=[assigned_rep_id]
    )


# ============================================================
# ربط المستخدم بالصيدلية
# ============================================================

class UserPharmacy(Base):

    __tablename__ = "user_pharmacies"

    id = Column(
        Integer,
        primary_key=True
    )

    user_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=False
    )

    pharmacy_id = Column(
        Integer,
        ForeignKey("pharmacies.id"),
        nullable=False
    )

    user = relationship("User")

    pharmacy = relationship("Pharmacy")

    __table_args__ = (
        UniqueConstraint(
            "user_id",
            "pharmacy_id",
            name="uq_user_pharmacy"
        ),
    )


# ============================================================
# الصيادلة
# ============================================================

class Pharmacist(Base):

    __tablename__ = "pharmacists"

    id = Column(
        Integer,
        primary_key=True
    )

    full_name = Column(
        String(255),
        nullable=False
    )

    phone = Column(
        String(50),
        nullable=True
    )

    pharmacy_id = Column(
        Integer,
        ForeignKey("pharmacies.id"),
        nullable=True
    )

    classification = Column(
        String(100),
        nullable=True
    )

    notes = Column(
        Text,
        nullable=True
    )

    pharmacy = relationship("Pharmacy")


# ============================================================
# المنتجات
# ============================================================

class Product(Base):

    __tablename__ = "products"

    id = Column(
        Integer,
        primary_key=True
    )

    code = Column(
        String(100),
        unique=True,
        nullable=False,
        index=True
    )

    brand_name = Column(
        String(255),
        nullable=False,
        index=True
    )

    scientific_name = Column(
        String(255),
        nullable=True,
        index=True
    )

    active_ingredient = Column(
        String(255),
        nullable=True
    )

    concentration = Column(
        String(100),
        nullable=True
    )

    dosage_form = Column(
        String(100),
        nullable=True
    )

    package_size = Column(
        String(100),
        nullable=True
    )

    therapeutic_class = Column(
        String(255),
        nullable=True
    )

    manufacturer = Column(
        String(255),
        nullable=True
    )

    country_of_origin = Column(
        String(100),
        nullable=True
    )

    price = Column(
        Float,
        default=0
    )

    minimum_stock = Column(
        Integer,
        default=0
    )

    scientific_notes = Column(
        Text,
        nullable=True
    )

    image_file_id = Column(
        String(255),
        nullable=True
    )

    leaflet_file_id = Column(
        String(255),
        nullable=True
    )

    is_active = Column(
        Boolean,
        default=True
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )


# ============================================================
# ربط المنتجات بالتخصصات
# ============================================================

class ProductSpecialty(Base):

    __tablename__ = "product_specialties"

    id = Column(
        Integer,
        primary_key=True
    )

    product_id = Column(
        Integer,
        ForeignKey("products.id"),
        nullable=False
    )

    specialty_id = Column(
        Integer,
        ForeignKey("specialties.id"),
        nullable=False
    )

    product = relationship("Product")

    specialty = relationship("Specialty")


# ============================================================
# التشغيلات
# ============================================================

class Batch(Base):

    __tablename__ = "batches"

    id = Column(
        Integer,
        primary_key=True
    )

    product_id = Column(
        Integer,
        ForeignKey("products.id"),
        nullable=False
    )

    batch_number = Column(
        String(100),
        nullable=False,
        index=True
    )

    manufacture_date = Column(
        Date,
        nullable=True
    )

    expiry_date = Column(
        Date,
        nullable=True,
        index=True
    )

    product = relationship("Product")

    __table_args__ = (
        UniqueConstraint(
            "product_id",
            "batch_number",
            name="uq_product_batch"
        ),
    )


# ============================================================
# المخازن
# ============================================================

class Warehouse(Base):

    __tablename__ = "warehouses"

    id = Column(
        Integer,
        primary_key=True
    )

    name = Column(
        String(255),
        nullable=False
    )

    governorate = Column(
        String(100),
        nullable=True
    )

    city = Column(
        String(100),
        nullable=True
    )

    address = Column(
        Text,
        nullable=True
    )

    manager_name = Column(
        String(255),
        nullable=True
    )

    phone = Column(
        String(50),
        nullable=True
    )

    is_active = Column(
        Boolean,
        default=True
    )


# ============================================================
# مخزون المخازن
# ============================================================

class WarehouseStock(Base):

    __tablename__ = "warehouse_stock"

    id = Column(
        Integer,
        primary_key=True
    )

    warehouse_id = Column(
        Integer,
        ForeignKey("warehouses.id"),
        nullable=False
    )

    product_id = Column(
        Integer,
        ForeignKey("products.id"),
        nullable=False
    )

    batch_id = Column(
        Integer,
        ForeignKey("batches.id"),
        nullable=True
    )

    quantity = Column(
        Integer,
        default=0
    )

    reserved_quantity = Column(
        Integer,
        default=0
    )

    updated_at = Column(
        DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow
    )

    warehouse = relationship("Warehouse")

    product = relationship("Product")

    batch = relationship("Batch")

    __table_args__ = (
        UniqueConstraint(
            "warehouse_id",
            "product_id",
            "batch_id",
            name="uq_warehouse_product_batch"
        ),
    )


# ============================================================
# مخزون الصيدليات
# ============================================================

class PharmacyStock(Base):

    __tablename__ = "pharmacy_stock"

    id = Column(
        Integer,
        primary_key=True
    )

    pharmacy_id = Column(
        Integer,
        ForeignKey("pharmacies.id"),
        nullable=False,
        index=True
    )

    product_id = Column(
        Integer,
        ForeignKey("products.id"),
        nullable=False
    )

    batch_id = Column(
        Integer,
        ForeignKey("batches.id"),
        nullable=True
    )

    quantity = Column(
        Integer,
        default=0
    )

    last_reported_by = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=True
    )

    updated_at = Column(
        DateTime,
        default=datetime.utcnow,
        onupdate=datetime.utcnow
    )

    pharmacy = relationship("Pharmacy")

    product = relationship("Product")

    batch = relationship("Batch")

    __table_args__ = (
        UniqueConstraint(
            "pharmacy_id",
            "product_id",
            "batch_id",
            name="uq_pharmacy_product_batch"
        ),
    )


# ============================================================
# حركات المخزون
# ============================================================

class StockMovement(Base):

    __tablename__ = "stock_movements"

    id = Column(
        Integer,
        primary_key=True
    )

    location_type = Column(
        String(50),
        nullable=False
    )

    location_id = Column(
        Integer,
        nullable=False
    )

    product_id = Column(
        Integer,
        ForeignKey("products.id"),
        nullable=False
    )

    batch_id = Column(
        Integer,
        ForeignKey("batches.id"),
        nullable=True
    )

    movement_type = Column(
        String(100),
        nullable=False
    )

    quantity_before = Column(
        Integer,
        default=0
    )

    quantity_change = Column(
        Integer,
        default=0
    )

    quantity_after = Column(
        Integer,
        default=0
    )

    user_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=True
    )

    notes = Column(
        Text,
        nullable=True
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )


# ============================================================
# زيارات الأطباء
# ============================================================

class DoctorVisit(Base):

    __tablename__ = "doctor_visits"

    id = Column(
        Integer,
        primary_key=True
    )

    doctor_id = Column(
        Integer,
        ForeignKey("doctors.id"),
        nullable=False
    )

    representative_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=True
    )

    visit_date = Column(
        Date,
        default=date.today
    )

    interest_level = Column(
        String(100),
        nullable=True
    )

    notes = Column(
        Text,
        nullable=True
    )

    next_visit_date = Column(
        Date,
        nullable=True
    )

    doctor = relationship("Doctor")

    representative = relationship("User")


# ============================================================
# منتجات زيارة الطبيب
# ============================================================

class DoctorVisitProduct(Base):

    __tablename__ = "doctor_visit_products"

    id = Column(
        Integer,
        primary_key=True
    )

    visit_id = Column(
        Integer,
        ForeignKey("doctor_visits.id"),
        nullable=False
    )

    product_id = Column(
        Integer,
        ForeignKey("products.id"),
        nullable=False
    )

    interest_score = Column(
        Float,
        default=0
    )


# ============================================================
# زيارات الصيدليات
# ============================================================

class PharmacyVisit(Base):

    __tablename__ = "pharmacy_visits"

    id = Column(
        Integer,
        primary_key=True
    )

    pharmacy_id = Column(
        Integer,
        ForeignKey("pharmacies.id"),
        nullable=False
    )

    representative_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=True
    )

    visit_date = Column(
        Date,
        default=date.today
    )

    notes = Column(
        Text,
        nullable=True
    )

    next_visit_date = Column(
        Date,
        nullable=True
    )


# ============================================================
# طلب زيارة
# ============================================================

class VisitRequest(Base):

    __tablename__ = "visit_requests"

    id = Column(
        Integer,
        primary_key=True
    )

    pharmacy_id = Column(
        Integer,
        ForeignKey("pharmacies.id"),
        nullable=False
    )

    requested_by_user_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=False
    )

    status = Column(
        String(50),
        default="pending"
    )

    notes = Column(
        Text,
        nullable=True
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )


# ============================================================
# التذكيرات
# ============================================================

class Reminder(Base):

    __tablename__ = "reminders"

    id = Column(
        Integer,
        primary_key=True
    )

    user_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=False
    )

    title = Column(
        String(255),
        nullable=False
    )

    message = Column(
        Text,
        nullable=True
    )

    reminder_date = Column(
        DateTime,
        nullable=False
    )

    is_sent = Column(
        Boolean,
        default=False
    )


# ============================================================
# سجل العمليات
# ============================================================

class AuditLog(Base):

    __tablename__ = "audit_logs"

    id = Column(
        Integer,
        primary_key=True
    )

    user_id = Column(
        Integer,
        ForeignKey("users.id"),
        nullable=True
    )

    action = Column(
        String(255),
        nullable=False
    )

    entity_type = Column(
        String(100),
        nullable=True
    )

    entity_id = Column(
        Integer,
        nullable=True
    )

    old_value = Column(
        Text,
        nullable=True
    )

    new_value = Column(
        Text,
        nullable=True
    )

    created_at = Column(
        DateTime,
        default=datetime.utcnow
    )


# ============================================================
# إنشاء الجداول
# ============================================================

def initialize_database():

    Base.metadata.create_all(
        bind=engine
    )


# ============================================================
# أدوات المستخدم
# ============================================================

def get_or_create_user(update: Update):

    telegram_user = update.effective_user

    if not telegram_user:
        return None

    db = SessionLocal()

    try:

        user = db.query(User).filter(
            User.telegram_id == telegram_user.id
        ).first()

        if not user:

            role = ROLE_REPRESENTATIVE

            if telegram_user.id in ADMIN_IDS:
                role = ROLE_ADMIN

            user = User(
                telegram_id=telegram_user.id,
                full_name=telegram_user.full_name,
                username=telegram_user.username,
                role=role
            )

            db.add(user)
            db.commit()
            db.refresh(user)

        else:

            user.full_name = telegram_user.full_name
            user.username = telegram_user.username

            db.commit()

        return {
            "id": user.id,
            "telegram_id": user.telegram_id,
            "full_name": user.full_name,
            "role": user.role
        }

    finally:

        db.close()


def is_admin(user_data):

    if not user_data:
        return False

    return user_data["role"] == ROLE_ADMIN


# ============================================================
# لوحة المدير
# ============================================================

def admin_keyboard():

    keyboard = [

        [
            InlineKeyboardButton(
                "👨‍⚕️ الأطباء",
                callback_data="admin_doctors"
            ),

            InlineKeyboardButton(
                "🏪 الصيدليات",
                callback_data="admin_pharmacies"
            ),
        ],

        [
            InlineKeyboardButton(
                "💊 المنتجات",
                callback_data="admin_products"
            ),

            InlineKeyboardButton(
                "📦 المخازن",
                callback_data="admin_warehouses"
            ),
        ],

        [
            InlineKeyboardButton(
                "🏥 المستشفيات",
                callback_data="admin_hospitals"
            ),

            InlineKeyboardButton(
                "👨‍🔬 المستخدمون",
                callback_data="admin_users"
            ),
        ],

        [
            InlineKeyboardButton(
                "📊 التقارير",
                callback_data="admin_reports"
            ),

            InlineKeyboardButton(
                "⚠️ التنبيهات",
                callback_data="admin_alerts"
            ),
        ],

        [
            InlineKeyboardButton(
                "🔎 البحث",
                callback_data="search"
            ),

            InlineKeyboardButton(
                "📈 لوحة الإحصائيات",
                callback_data="dashboard"
            ),
        ],

    ]

    return InlineKeyboardMarkup(keyboard)


# ============================================================
# لوحة المندوب
# ============================================================

def representative_keyboard():

    keyboard = [

        [
            InlineKeyboardButton(
                "👨‍⚕️ الأطباء",
                callback_data="rep_doctors"
            ),

            InlineKeyboardButton(
                "🏪 الصيدليات",
                callback_data="rep_pharmacies"
            ),
        ],

        [
            InlineKeyboardButton(
                "🗓️ زيارات اليوم",
                callback_data="today_visits"
            ),

            InlineKeyboardButton(
                "➕ تسجيل زيارة",
                callback_data="add_visit"
            ),
        ],

        [
            InlineKeyboardButton(
                "💊 المنتجات",
                callback_data="products_list"
            ),

            InlineKeyboardButton(
                "🔎 البحث",
                callback_data="search"
            ),
        ],

        [
            InlineKeyboardButton(
                "🔔 التذكيرات",
                callback_data="my_reminders"
            ),
        ]

    ]

    return InlineKeyboardMarkup(keyboard)


# ============================================================
# لوحة صاحب الصيدلية
# ============================================================

def pharmacy_owner_keyboard():

    keyboard = [

        [
            InlineKeyboardButton(
                "🏪 صيدليتي",
                callback_data="my_pharmacy"
            )
        ],

        [
            InlineKeyboardButton(
                "📦 مخزوني",
                callback_data="pharmacy_inventory"
            ),

            InlineKeyboardButton(
                "➕ إضافة صنف",
                callback_data="pharmacy_add_stock"
            ),
        ],

        [
            InlineKeyboardButton(
                "✏️ تحديث كمية",
                callback_data="pharmacy_update_stock"
            ),

            InlineKeyboardButton(
                "📋 قائمة الأصناف",
                callback_data="pharmacy_products"
            ),
        ],

        [
            InlineKeyboardButton(
                "⚠️ مخزون منخفض",
                callback_data="pharmacy_low_stock"
            ),

            InlineKeyboardButton(
                "⏳ قرب الانتهاء",
                callback_data="pharmacy_expiry"
            ),
        ],

        [
            InlineKeyboardButton(
                "🗓️ طلب زيارة مندوب",
                callback_data="request_visit"
            ),

            InlineKeyboardButton(
                "💬 إرسال ملاحظة",
                callback_data="pharmacy_note"
            ),
        ]

    ]

    return InlineKeyboardMarkup(keyboard)


# ============================================================
# لوحة حسب الصلاحية
# ============================================================

def get_main_keyboard(user):

    if user["role"] == ROLE_ADMIN:
        return admin_keyboard()

    if user["role"] in [
        ROLE_PHARMACY_OWNER,
        ROLE_PHARMACIST
    ]:
        return pharmacy_owner_keyboard()

    return representative_keyboard()


# ============================================================
# حالة المحادثات
# ============================================================

(
    ADD_DOCTOR_NAME,
    ADD_DOCTOR_PHONE,
    ADD_DOCTOR_SPECIALTY,
    ADD_DOCTOR_CATEGORY,
    ADD_DOCTOR_CITY,

    ADD_PHARMACY_NAME,
    ADD_PHARMACY_OWNER,
    ADD_PHARMACY_PHONE,
    ADD_PHARMACY_CITY,

    ADD_PRODUCT_CODE,
    ADD_PRODUCT_NAME,
    ADD_PRODUCT_SCIENTIFIC,
    ADD_PRODUCT_CONCENTRATION,
    ADD_PRODUCT_FORM,

    ADD_STOCK_PRODUCT,
    ADD_STOCK_BATCH,
    ADD_STOCK_QUANTITY,
    ADD_STOCK_EXPIRY,

    SEARCH_TEXT,

) = range(19)


# ============================================================
# START
# ============================================================

async def start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user = get_or_create_user(update)

    text = (
        f"🩺 أهلاً بك في نظام {COMPANY_NAME}\n\n"
        "نظام إدارة المندوبين العلميين "
        "والأطباء والصيدليات والمخزون."
    )

    await update.message.reply_text(
        text,
        reply_markup=get_main_keyboard(user)
    )


# ============================================================
# لوحة الإحصائيات
# ============================================================

async def dashboard_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    db = SessionLocal()

    try:

        doctors_count = db.query(
            func.count(Doctor.id)
        ).scalar() or 0

        pharmacies_count = db.query(
            func.count(Pharmacy.id)
        ).scalar() or 0

        pharmacists_count = db.query(
            func.count(Pharmacist.id)
        ).scalar() or 0

        products_count = db.query(
            func.count(Product.id)
        ).scalar() or 0

        batches_count = db.query(
            func.count(Batch.id)
        ).scalar() or 0

        hospitals_count = db.query(
            func.count(Hospital.id)
        ).scalar() or 0

        today = date.today()

        visits_today = db.query(
            func.count(DoctorVisit.id)
        ).filter(
            DoctorVisit.visit_date == today
        ).scalar() or 0

        text = (
            "📈 لوحة التحكم\n\n"
            f"👨‍⚕️ الأطباء: {doctors_count}\n"
            f"🏪 الصيدليات: {pharmacies_count}\n"
            f"💊 الصيادلة: {pharmacists_count}\n"
            f"💊 المنتجات: {products_count}\n"
            f"🧪 التشغيلات: {batches_count}\n"
            f"🏥 المستشفيات: {hospitals_count}\n"
            f"🗓️ زيارات اليوم: {visits_today}"
        )

        await query.message.reply_text(text)

    finally:

        db.close()


# ============================================================
# إضافة طبيب
# ============================================================

async def add_doctor_start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    await query.message.reply_text(
        "👨‍⚕️ أدخل الاسم الكامل للطبيب:"
    )

    return ADD_DOCTOR_NAME


async def add_doctor_name(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["doctor_name"] = (
        update.message.text.strip()
    )

    await update.message.reply_text(
        "📞 أدخل رقم هاتف الطبيب:"
    )

    return ADD_DOCTOR_PHONE


async def add_doctor_phone(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["doctor_phone"] = (
        update.message.text.strip()
    )

    db = SessionLocal()

    try:

        specialties = db.query(
            Specialty
        ).limit(20).all()

        if specialties:

            text = "🩺 أدخل اسم تخصص الطبيب:\n\n"

            for specialty in specialties:
                text += f"• {specialty.name}\n"

        else:

            text = (
                "🩺 أدخل تخصص الطبيب.\n"
                "مثال: أطفال"
            )

        await update.message.reply_text(text)

    finally:

        db.close()

    return ADD_DOCTOR_SPECIALTY


async def add_doctor_specialty(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    specialty_name = update.message.text.strip()

    db = SessionLocal()

    try:

        specialty = db.query(
            Specialty
        ).filter(
            Specialty.name == specialty_name
        ).first()

        if not specialty:

            specialty = Specialty(
                name=specialty_name
            )

            db.add(specialty)
            db.commit()
            db.refresh(specialty)

        context.user_data["specialty_id"] = (
            specialty.id
        )

    finally:

        db.close()

    await update.message.reply_text(
        "⭐ اختر تصنيف الطبيب:\n"
        "A / B / C / D"
    )

    return ADD_DOCTOR_CATEGORY


async def add_doctor_category(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    category = (
        update.message.text.strip().upper()
    )

    if category not in ["A", "B", "C", "D"]:

        await update.message.reply_text(
            "❌ التصنيف غير صحيح.\n"
            "أدخل A أو B أو C أو D."
        )

        return ADD_DOCTOR_CATEGORY

    context.user_data["doctor_category"] = category

    await update.message.reply_text(
        "📍 أدخل المدينة أو المنطقة:"
    )

    return ADD_DOCTOR_CITY


async def add_doctor_city(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    city = update.message.text.strip()

    user_data = get_or_create_user(update)

    db = SessionLocal()

    try:

        doctor = Doctor(

            full_name=context.user_data.get(
                "doctor_name"
            ),

            phone=context.user_data.get(
                "doctor_phone"
            ),

            specialty_id=context.user_data.get(
                "specialty_id"
            ),

            category=context.user_data.get(
                "doctor_category"
            ),

            city=city,

            assigned_rep_id=user_data["id"]

        )

        db.add(doctor)

        db.commit()

        db.refresh(doctor)

        add_audit_log(
            db=db,
            user_id=user_data["id"],
            action="إضافة طبيب",
            entity_type="doctor",
            entity_id=doctor.id
        )

        db.commit()

        await update.message.reply_text(
            f"✅ تم إضافة الطبيب بنجاح.\n\n"
            f"👨‍⚕️ الاسم: {doctor.full_name}\n"
            f"🆔 الرقم: {doctor.id}\n"
            f"⭐ التصنيف: {doctor.category}"
        )

    finally:

        db.close()

    context.user_data.clear()

    return ConversationHandler.END


# ============================================================
# إضافة صيدلية
# ============================================================

async def add_pharmacy_start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    await query.message.reply_text(
        "🏪 أدخل اسم الصيدلية:"
    )

    return ADD_PHARMACY_NAME


async def add_pharmacy_name(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["pharmacy_name"] = (
        update.message.text.strip()
    )

    await update.message.reply_text(
        "👤 أدخل اسم صاحب الصيدلية:"
    )

    return ADD_PHARMACY_OWNER


async def add_pharmacy_owner(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["pharmacy_owner"] = (
        update.message.text.strip()
    )

    await update.message.reply_text(
        "📞 أدخل رقم الهاتف:"
    )

    return ADD_PHARMACY_PHONE


async def add_pharmacy_phone(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["pharmacy_phone"] = (
        update.message.text.strip()
    )

    await update.message.reply_text(
        "📍 أدخل المدينة:"
    )

    return ADD_PHARMACY_CITY


async def add_pharmacy_city(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_data = get_or_create_user(update)

    db = SessionLocal()

    try:

        pharmacy = Pharmacy(

            name=context.user_data.get(
                "pharmacy_name"
            ),

            owner_name=context.user_data.get(
                "pharmacy_owner"
            ),

            phone=context.user_data.get(
                "pharmacy_phone"
            ),

            city=update.message.text.strip(),

            assigned_rep_id=user_data["id"]

        )

        db.add(pharmacy)

        db.commit()

        db.refresh(pharmacy)

        add_audit_log(
            db=db,
            user_id=user_data["id"],
            action="إضافة صيدلية",
            entity_type="pharmacy",
            entity_id=pharmacy.id
        )

        db.commit()

        await update.message.reply_text(
            f"✅ تم إضافة الصيدلية بنجاح.\n\n"
            f"🏪 {pharmacy.name}\n"
            f"🆔 رقم الصيدلية: {pharmacy.id}"
        )

    finally:

        db.close()

    context.user_data.clear()

    return ConversationHandler.END


# ============================================================
# إضافة منتج
# ============================================================

async def add_product_start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    await query.message.reply_text(
        "🔢 أدخل كود المنتج:"
    )

    return ADD_PRODUCT_CODE


async def add_product_code(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["product_code"] = (
        update.message.text.strip()
    )

    await update.message.reply_text(
        "💊 أدخل الاسم التجاري:"
    )

    return ADD_PRODUCT_NAME


async def add_product_name(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["product_name"] = (
        update.message.text.strip()
    )

    await update.message.reply_text(
        "🔬 أدخل الاسم العلمي أو المادة الفعالة:"
    )

    return ADD_PRODUCT_SCIENTIFIC


async def add_product_scientific(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["product_scientific"] = (
        update.message.text.strip()
    )

    await update.message.reply_text(
        "⚗️ أدخل التركيز.\nمثال: 500mg"
    )

    return ADD_PRODUCT_CONCENTRATION


async def add_product_concentration(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data["product_concentration"] = (
        update.message.text.strip()
    )

    await update.message.reply_text(
        "💉 أدخل الشكل الدوائي.\n"
        "مثال: أقراص / شراب / حقن"
    )

    return ADD_PRODUCT_FORM


async def add_product_form(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_data = get_or_create_user(update)

    db = SessionLocal()

    try:

        existing = db.query(Product).filter(
            Product.code ==
            context.user_data.get(
                "product_code"
            )
        ).first()

        if existing:

            await update.message.reply_text(
                "❌ هذا الكود مستخدم مسبقاً."
            )

            return ConversationHandler.END

        product = Product(

            code=context.user_data.get(
                "product_code"
            ),

            brand_name=context.user_data.get(
                "product_name"
            ),

            scientific_name=context.user_data.get(
                "product_scientific"
            ),

            active_ingredient=context.user_data.get(
                "product_scientific"
            ),

            concentration=context.user_data.get(
                "product_concentration"
            ),

            dosage_form=update.message.text.strip()

        )

        db.add(product)

        db.commit()

        db.refresh(product)

        add_audit_log(
            db=db,
            user_id=user_data["id"],
            action="إضافة منتج",
            entity_type="product",
            entity_id=product.id
        )

        db.commit()

        await update.message.reply_text(
            f"✅ تم إضافة المنتج.\n\n"
            f"💊 {product.brand_name}\n"
            f"🆔 {product.id}\n"
            f"🔢 {product.code}"
        )

    finally:

        db.close()

    context.user_data.clear()

    return ConversationHandler.END


# ============================================================
# اختيار صيدلية المستخدم
# ============================================================

def get_user_pharmacies(
    user_id
):

    db = SessionLocal()

    try:

        links = db.query(
            UserPharmacy
        ).filter(
            UserPharmacy.user_id == user_id
        ).all()

        pharmacies = []

        for link in links:

            pharmacy = db.query(
                Pharmacy
            ).filter(
                Pharmacy.id == link.pharmacy_id
            ).first()

            if pharmacy:
                pharmacies.append(pharmacy)

        return pharmacies

    finally:

        db.close()


# ============================================================
# عرض صيدلية المستخدم
# ============================================================

async def my_pharmacy_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    user_data = get_or_create_user(update)

    pharmacies = get_user_pharmacies(
        user_data["id"]
    )

    if not pharmacies:

        await query.message.reply_text(
            "⚠️ لم يتم ربط حسابك بأي صيدلية بعد.\n\n"
            "يرجى أن يقوم المدير بربط حسابك بالصيدلية."
        )

        return

    text = "🏪 صيدلياتي:\n\n"

    for pharmacy in pharmacies:

        text += (
            f"🆔 {pharmacy.id}\n"
            f"🏪 {pharmacy.name}\n"
            f"👤 {pharmacy.owner_name or '-'}\n"
            f"📞 {pharmacy.phone or '-'}\n"
            f"📍 {pharmacy.city or '-'}\n\n"
        )

    await query.message.reply_text(text)


# ============================================================
# إضافة مخزون الصيدلية
# ============================================================

async def pharmacy_add_stock_start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    user_data = get_or_create_user(update)

    pharmacies = get_user_pharmacies(
        user_data["id"]
    )

    if not pharmacies:

        await query.message.reply_text(
            "❌ لا توجد صيدلية مرتبطة بحسابك."
        )

        return ConversationHandler.END

    context.user_data[
        "stock_pharmacy_id"
    ] = pharmacies[0].id

    await query.message.reply_text(
        "💊 أدخل اسم المنتج أو كود المنتج:"
    )

    return ADD_STOCK_PRODUCT


async def pharmacy_add_stock_product(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    search = update.message.text.strip()

    db = SessionLocal()

    try:

        product = db.query(Product).filter(
            (Product.code == search)
            |
            (Product.brand_name.ilike(
                f"%{search}%"
            ))
        ).first()

        if not product:

            await update.message.reply_text(
                "❌ المنتج غير موجود في قاعدة الشركة.\n"
                "أدخل الاسم أو الكود الصحيح."
            )

            return ADD_STOCK_PRODUCT

        context.user_data[
            "stock_product_id"
        ] = product.id

        await update.message.reply_text(
            f"💊 تم اختيار: {product.brand_name}\n\n"
            "🧪 أدخل رقم التشغيلة Batch Number:"
        )

    finally:

        db.close()

    return ADD_STOCK_BATCH


async def pharmacy_add_stock_batch(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data[
        "stock_batch_number"
    ] = update.message.text.strip()

    await update.message.reply_text(
        "🔢 أدخل الكمية الموجودة:"
    )

    return ADD_STOCK_QUANTITY


async def pharmacy_add_stock_quantity(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    try:

        quantity = int(
            update.message.text.strip()
        )

        if quantity < 0:
            raise ValueError

    except ValueError:

        await update.message.reply_text(
            "❌ أدخل كمية صحيحة."
        )

        return ADD_STOCK_QUANTITY

    context.user_data[
        "stock_quantity"
    ] = quantity

    await update.message.reply_text(
        "📅 أدخل تاريخ الانتهاء بهذا الشكل:\n"
        "YYYY-MM-DD\n\n"
        "مثال:\n"
        "2028-12-31"
    )

    return ADD_STOCK_EXPIRY


async def pharmacy_add_stock_expiry(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    try:

        expiry = datetime.strptime(
            update.message.text.strip(),
            "%Y-%m-%d"
        ).date()

    except ValueError:

        await update.message.reply_text(
            "❌ التاريخ غير صحيح.\n"
            "استخدم YYYY-MM-DD"
        )

        return ADD_STOCK_EXPIRY

    user_data = get_or_create_user(update)

    db = SessionLocal()

    try:

        product_id = context.user_data.get(
            "stock_product_id"
        )

        pharmacy_id = context.user_data.get(
            "stock_pharmacy_id"
        )

        batch_number = context.user_data.get(
            "stock_batch_number"
        )

        quantity = context.user_data.get(
            "stock_quantity"
        )

        batch = db.query(Batch).filter(
            Batch.product_id == product_id,
            Batch.batch_number == batch_number
        ).first()

        if not batch:

            batch = Batch(
                product_id=product_id,
                batch_number=batch_number,
                expiry_date=expiry
            )

            db.add(batch)
            db.commit()
            db.refresh(batch)

        else:

            if not batch.expiry_date:
                batch.expiry_date = expiry

        stock = db.query(
            PharmacyStock
        ).filter(
            PharmacyStock.pharmacy_id == pharmacy_id,
            PharmacyStock.product_id == product_id,
            PharmacyStock.batch_id == batch.id
        ).first()

        if stock:

            before = stock.quantity

            stock.quantity += quantity

            after = stock.quantity

        else:

            before = 0

            after = quantity

            stock = PharmacyStock(
                pharmacy_id=pharmacy_id,
                product_id=product_id,
                batch_id=batch.id,
                quantity=quantity,
                last_reported_by=user_data["id"]
            )

            db.add(stock)

        movement = StockMovement(

            location_type="pharmacy",

            location_id=pharmacy_id,

            product_id=product_id,

            batch_id=batch.id,

            movement_type="إضافة مخزون",

            quantity_before=before,

            quantity_change=quantity,

            quantity_after=after,

            user_id=user_data["id"],

            notes="تم التحديث بواسطة الصيدلية"

        )

        db.add(movement)

        add_audit_log(
            db=db,
            user_id=user_data["id"],
            action="تحديث مخزون الصيدلية",
            entity_type="pharmacy_stock",
            entity_id=stock.id,
            old_value=str(before),
            new_value=str(after)
        )

        db.commit()

        product = db.query(Product).filter(
            Product.id == product_id
        ).first()

        await update.message.reply_text(
            "✅ تم تحديث مخزون الصيدلية بنجاح.\n\n"
            f"💊 المنتج: {product.brand_name}\n"
            f"🧪 التشغيلة: {batch.batch_number}\n"
            f"🔢 الكمية الحالية: {after}\n"
            f"📅 الانتهاء: {expiry}"
        )

    except Exception as error:

        db.rollback()

        logger.exception(error)

        await update.message.reply_text(
            "❌ حدث خطأ أثناء حفظ المخزون."
        )

    finally:

        db.close()

    context.user_data.clear()

    return ConversationHandler.END


# ============================================================
# عرض مخزون الصيدلية
# ============================================================

async def pharmacy_inventory_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    user_data = get_or_create_user(update)

    pharmacies = get_user_pharmacies(
        user_data["id"]
    )

    if not pharmacies:

        await query.message.reply_text(
            "❌ لا توجد صيدلية مرتبطة بحسابك."
        )

        return

    db = SessionLocal()

    try:

        text = "📦 مخزون صيدليتك:\n\n"

        count = 0

        for pharmacy in pharmacies:

            stocks = db.query(
                PharmacyStock
            ).filter(
                PharmacyStock.pharmacy_id ==
                pharmacy.id
            ).all()

            text += (
                f"🏪 {pharmacy.name}\n"
                "────────────────\n"
            )

            for stock in stocks:

                product = db.query(
                    Product
                ).filter(
                    Product.id == stock.product_id
                ).first()

                batch = None

                if stock.batch_id:

                    batch = db.query(
                        Batch
                    ).filter(
                        Batch.id == stock.batch_id
                    ).first()

                if product:

                    expiry_text = "-"

                    if batch and batch.expiry_date:
                        expiry_text = str(
                            batch.expiry_date
                        )

                    text += (
                        f"💊 {product.brand_name}\n"
                        f"🔢 الكمية: {stock.quantity}\n"
                        f"🧪 التشغيلة: "
                        f"{batch.batch_number if batch else '-'}\n"
                        f"📅 الانتهاء: "
                        f"{expiry_text}\n\n"
                    )

                    count += 1

                    if len(text) > 3500:

                        await query.message.reply_text(
                            text
                        )

                        text = ""

        if count == 0:

            text += "لا توجد أصناف مسجلة."

        if text:

            await query.message.reply_text(text)

    finally:

        db.close()


# ============================================================
# قائمة المنتجات
# ============================================================

async def products_list_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    db = SessionLocal()

    try:

        products = db.query(Product).filter(
            Product.is_active == True
        ).order_by(
            Product.brand_name
        ).limit(100).all()

        if not products:

            await query.message.reply_text(
                "لا توجد منتجات حالياً."
            )

            return

        text = "💊 قائمة المنتجات:\n\n"

        for product in products:

            text += (
                f"🆔 {product.id}\n"
                f"💊 {product.brand_name}\n"
                f"🔬 {product.scientific_name or '-'}\n"
                f"⚗️ {product.concentration or '-'}\n"
                f"💉 {product.dosage_form or '-'}\n\n"
            )

            if len(text) > 3500:

                await query.message.reply_text(
                    text
                )

                text = ""

        if text:

            await query.message.reply_text(text)

    finally:

        db.close()


# ============================================================
# البحث
# ============================================================

async def search_start(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    await query.message.reply_text(
        "🔎 أدخل ما تريد البحث عنه:\n\n"
        "• اسم طبيب\n"
        "• اسم صيدلية\n"
        "• اسم منتج\n"
        "• اسم علمي\n"
        "• رقم هاتف\n"
        "• رقم تشغيلة"
    )

    return SEARCH_TEXT


async def search_execute(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    search = update.message.text.strip()

    db = SessionLocal()

    try:

        text = f"🔎 نتائج البحث عن: {search}\n\n"

        doctors = db.query(Doctor).filter(
            Doctor.full_name.ilike(
                f"%{search}%"
            )
        ).limit(10).all()

        if doctors:

            text += "👨‍⚕️ الأطباء:\n"

            for doctor in doctors:

                text += (
                    f"• {doctor.full_name}\n"
                    f"  🩺 {doctor.category}\n"
                    f"  📞 {doctor.phone or '-'}\n"
                )

            text += "\n"

        pharmacies = db.query(
            Pharmacy
        ).filter(
            Pharmacy.name.ilike(
                f"%{search}%"
            )
        ).limit(10).all()

        if pharmacies:

            text += "🏪 الصيدليات:\n"

            for pharmacy in pharmacies:

                text += (
                    f"• {pharmacy.name}\n"
                    f"  📞 {pharmacy.phone or '-'}\n"
                    f"  📍 {pharmacy.city or '-'}\n"
                )

            text += "\n"

        products = db.query(Product).filter(
            (
                Product.brand_name.ilike(
                    f"%{search}%"
                )
            )
            |
            (
                Product.scientific_name.ilike(
                    f"%{search}%"
                )
            )
            |
            (
                Product.code.ilike(
                    f"%{search}%"
                )
            )
        ).limit(10).all()

        if products:

            text += "💊 المنتجات:\n"

            for product in products:

                text += (
                    f"• {product.brand_name}\n"
                    f"  🔢 {product.code}\n"
                    f"  🔬 {product.scientific_name or '-'}\n"
                )

        if (
            not doctors
            and not pharmacies
            and not products
        ):

            text += "❌ لم يتم العثور على نتائج."

        await update.message.reply_text(text)

    finally:

        db.close()

    return ConversationHandler.END


# ============================================================
# التنبيهات
# ============================================================

def get_expiry_alerts():

    db = SessionLocal()

    try:

        today = date.today()

        warning_date = (
            today + timedelta(days=180)
        )

        batches = db.query(Batch).filter(
            Batch.expiry_date != None,
            Batch.expiry_date <= warning_date
        ).all()

        return batches

    finally:

        db.close()


async def alerts_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    db = SessionLocal()

    try:

        today = date.today()

        three_months = (
            today + timedelta(days=90)
        )

        one_month = (
            today + timedelta(days=30)
        )

        batches = db.query(Batch).filter(
            Batch.expiry_date != None,
            Batch.expiry_date <= (
                today + timedelta(days=180)
            )
        ).all()

        if not batches:

            await query.message.reply_text(
                "🟢 لا توجد تشغيلات قريبة من الانتهاء."
            )

            return

        text = "⚠️ تنبيهات الصلاحية:\n\n"

        for batch in batches:

            product = db.query(Product).filter(
                Product.id == batch.product_id
            ).first()

            if batch.expiry_date < today:
                status = "⚫ منتهي"

            elif batch.expiry_date <= one_month:
                status = "🔴 أقل من شهر"

            elif batch.expiry_date <= three_months:
                status = "🟠 أقل من 3 أشهر"

            else:
                status = "🟡 أقل من 6 أشهر"

            text += (
                f"{status}\n"
                f"💊 {product.brand_name if product else '-'}\n"
                f"🧪 {batch.batch_number}\n"
                f"📅 {batch.expiry_date}\n\n"
            )

        await query.message.reply_text(text)

    finally:

        db.close()


# ============================================================
# تقرير إجمالي المنتج
# ============================================================

def calculate_product_distribution(
    product_id
):

    db = SessionLocal()

    try:

        warehouse_total = db.query(
            func.coalesce(
                func.sum(
                    WarehouseStock.quantity
                ),
                0
            )
        ).filter(
            WarehouseStock.product_id ==
            product_id
        ).scalar()

        pharmacy_total = db.query(
            func.coalesce(
                func.sum(
                    PharmacyStock.quantity
                ),
                0
            )
        ).filter(
            PharmacyStock.product_id ==
            product_id
        ).scalar()

        pharmacy_count = db.query(
            func.count(
                func.distinct(
                    PharmacyStock.pharmacy_id
                )
            )
        ).filter(
            PharmacyStock.product_id ==
            product_id,
            PharmacyStock.quantity > 0
        ).scalar()

        return {

            "warehouse_total":
                warehouse_total or 0,

            "pharmacy_total":
                pharmacy_total or 0,

            "pharmacy_count":
                pharmacy_count or 0

        }

    finally:

        db.close()


# ============================================================
# التقارير
# ============================================================

async def reports_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    db = SessionLocal()

    try:

        total_doctors = db.query(
            func.count(Doctor.id)
        ).scalar() or 0

        category_a = db.query(
            func.count(Doctor.id)
        ).filter(
            Doctor.category == "A"
        ).scalar() or 0

        total_pharmacies = db.query(
            func.count(Pharmacy.id)
        ).scalar() or 0

        total_products = db.query(
            func.count(Product.id)
        ).scalar() or 0

        low_stock_products = 0

        products = db.query(Product).all()

        for product in products:

            total = db.query(
                func.coalesce(
                    func.sum(
                        WarehouseStock.quantity
                    ),
                    0
                )
            ).filter(
                WarehouseStock.product_id ==
                product.id
            ).scalar()

            if (
                product.minimum_stock > 0
                and total < product.minimum_stock
            ):
                low_stock_products += 1

        text = (
            "📊 التقرير العام\n\n"
            f"👨‍⚕️ إجمالي الأطباء: {total_doctors}\n"
            f"⭐ أطباء التصنيف A: {category_a}\n"
            f"🏪 إجمالي الصيدليات: {total_pharmacies}\n"
            f"💊 إجمالي المنتجات: {total_products}\n"
            f"📉 منتجات منخفضة المخزون: "
            f"{low_stock_products}"
        )

        await query.message.reply_text(text)

    finally:

        db.close()


# ============================================================
# طلب زيارة مندوب
# ============================================================

async def request_visit_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    user_data = get_or_create_user(update)

    pharmacies = get_user_pharmacies(
        user_data["id"]
    )

    if not pharmacies:

        await query.message.reply_text(
            "❌ لا توجد صيدلية مرتبطة بحسابك."
        )

        return

    db = SessionLocal()

    try:

        pharmacy = pharmacies[0]

        request = VisitRequest(

            pharmacy_id=pharmacy.id,

            requested_by_user_id=user_data["id"],

            status="pending",

            notes="طلب زيارة من صاحب الصيدلية"

        )

        db.add(request)

        db.commit()

        await query.message.reply_text(
            "✅ تم إرسال طلب زيارة المندوب.\n"
            "سيظهر الطلب للإدارة."
        )

    finally:

        db.close()


# ============================================================
# سجل العمليات
# ============================================================

def add_audit_log(
    db,
    user_id,
    action,
    entity_type=None,
    entity_id=None,
    old_value=None,
    new_value=None
):

    log = AuditLog(

        user_id=user_id,

        action=action,

        entity_type=entity_type,

        entity_id=entity_id,

        old_value=old_value,

        new_value=new_value

    )

    db.add(log)


# ============================================================
# عرض الأطباء
# ============================================================

async def doctors_list_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    db = SessionLocal()

    try:

        doctors = db.query(
            Doctor
        ).order_by(
            Doctor.full_name
        ).limit(50).all()

        if not doctors:

            await query.message.reply_text(
                "لا يوجد أطباء مسجلون."
            )

            return

        text = "👨‍⚕️ قائمة الأطباء:\n\n"

        for doctor in doctors:

            specialty_name = "-"

            if doctor.specialty:
                specialty_name = doctor.specialty.name

            text += (
                f"🆔 {doctor.id}\n"
                f"👨‍⚕️ {doctor.full_name}\n"
                f"🩺 {specialty_name}\n"
                f"⭐ التصنيف: {doctor.category}\n"
                f"📞 {doctor.phone or '-'}\n"
                f"📍 {doctor.city or '-'}\n\n"
            )

            if len(text) > 3500:

                await query.message.reply_text(
                    text
                )

                text = ""

        if text:
            await query.message.reply_text(text)

    finally:

        db.close()


# ============================================================
# عرض الصيدليات
# ============================================================

async def pharmacies_list_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    await query.answer()

    db = SessionLocal()

    try:

        pharmacies = db.query(
            Pharmacy
        ).order_by(
            Pharmacy.name
        ).limit(50).all()

        if not pharmacies:

            await query.message.reply_text(
                "لا توجد صيدليات مسجلة."
            )

            return

        text = "🏪 قائمة الصيدليات:\n\n"

        for pharmacy in pharmacies:

            text += (
                f"🆔 {pharmacy.id}\n"
                f"🏪 {pharmacy.name}\n"
                f"👤 {pharmacy.owner_name or '-'}\n"
                f"📞 {pharmacy.phone or '-'}\n"
                f"📍 {pharmacy.city or '-'}\n\n"
            )

        await query.message.reply_text(text)

    finally:

        db.close()


# ============================================================
# إضافة تخصصات افتراضية
# ============================================================

def seed_specialties():

    specialties = [

        "طب عام",

        "الباطنية",

        "الأطفال",

        "النساء والولادة",

        "القلب",

        "الجراحة",

        "العظام",

        "الأنف والأذن والحنجرة",

        "الجلدية",

        "العيون",

        "المسالك البولية",

        "الأعصاب",

        "الأورام",

        "الأسنان",

        "الغدد الصماء",

        "الكلى",

        "الجهاز الهضمي",

        "الصدرية",

        "الطوارئ",

        "التخدير",

        "الأشعة",

        "المختبرات"

    ]

    db = SessionLocal()

    try:

        for name in specialties:

            existing = db.query(
                Specialty
            ).filter(
                Specialty.name == name
            ).first()

            if not existing:

                db.add(
                    Specialty(
                        name=name
                    )
                )

        db.commit()

    finally:

        db.close()


# ============================================================
# معالجة الأزرار
# ============================================================

async def menu_callback(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    query = update.callback_query

    data = query.data

    if data == "admin_doctors":

        await doctors_list_callback(
            update,
            context
        )

    elif data == "admin_pharmacies":

        await pharmacies_list_callback(
            update,
            context
        )

    elif data == "rep_doctors":

        await doctors_list_callback(
            update,
            context
        )

    elif data == "rep_pharmacies":

        await pharmacies_list_callback(
            update,
            context
        )

    elif data == "products_list":

        await products_list_callback(
            update,
            context
        )

    elif data == "admin_alerts":

        await alerts_callback(
            update,
            context
        )

    elif data == "admin_reports":

        await reports_callback(
            update,
            context
        )

    else:

        await query.answer()

        await query.message.reply_text(
            "⚠️ هذه الخاصية ستكون متاحة في القائمة الكاملة."
        )


# ============================================================
# إلغاء المحادثة
# ============================================================

async def cancel(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    context.user_data.clear()

    await update.message.reply_text(
        "❌ تم إلغاء العملية."
    )

    return ConversationHandler.END


# ============================================================
# الأوامر الإدارية
# ============================================================

async def add_pharmacy_user_command(
    update: Update,
    context: ContextTypes.DEFAULT_TYPE
):

    user_data = get_or_create_user(update)

    if not is_admin(user_data):

        await update.message.reply_text(
            "❌ ليس لديك صلاحية."
        )

        return

    await update.message.reply_text(
        "لاستخدام الربط استخدم لاحقاً لوحة الإدارة.\n\n"
        "صيغة قاعدة البيانات:\n"
        "ربط user_id مع pharmacy_id."
    )


# ============================================================
# معالجة الأخطاء
# ============================================================

async def error_handler(
    update: object,
    context: ContextTypes.DEFAULT_TYPE
):

    logger.exception(
        "حدث خطأ:",
        exc_info=context.error
    )


# ============================================================
# بناء التطبيق
# ============================================================

def main():

    initialize_database()

    seed_specialties()

    application = Application.builder().token(
        BOT_TOKEN
    ).build()

    # ========================================================
    # START
    # ========================================================

    application.add_handler(

        CommandHandler(
            "start",
            start
        )

    )

    # ========================================================
    # إضافة طبيب
    # ========================================================

    doctor_conversation = ConversationHandler(

        entry_points=[

            CallbackQueryHandler(
                add_doctor_start,
                pattern="^add_doctor$"
            )

        ],

        states={

            ADD_DOCTOR_NAME: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_doctor_name
                )

            ],

            ADD_DOCTOR_PHONE: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_doctor_phone
                )

            ],

            ADD_DOCTOR_SPECIALTY: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_doctor_specialty
                )

            ],

            ADD_DOCTOR_CATEGORY: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_doctor_category
                )

            ],

            ADD_DOCTOR_CITY: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_doctor_city
                )

            ],

        },

        fallbacks=[

            CommandHandler(
                "cancel",
                cancel
            )

        ]

    )

    application.add_handler(
        doctor_conversation
    )

    # ========================================================
    # إضافة صيدلية
    # ========================================================

    pharmacy_conversation = ConversationHandler(

        entry_points=[

            CallbackQueryHandler(
                add_pharmacy_start,
                pattern="^add_pharmacy$"
            )

        ],

        states={

            ADD_PHARMACY_NAME: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_pharmacy_name
                )

            ],

            ADD_PHARMACY_OWNER: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_pharmacy_owner
                )

            ],

            ADD_PHARMACY_PHONE: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_pharmacy_phone
                )

            ],

            ADD_PHARMACY_CITY: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_pharmacy_city
                )

            ]

        },

        fallbacks=[

            CommandHandler(
                "cancel",
                cancel
            )

        ]

    )

    application.add_handler(
        pharmacy_conversation
    )

    # ========================================================
    # إضافة منتج
    # ========================================================

    product_conversation = ConversationHandler(

        entry_points=[

            CallbackQueryHandler(
                add_product_start,
                pattern="^add_product$"
            )

        ],

        states={

            ADD_PRODUCT_CODE: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_product_code
                )

            ],

            ADD_PRODUCT_NAME: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_product_name
                )

            ],

            ADD_PRODUCT_SCIENTIFIC: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_product_scientific
                )

            ],

            ADD_PRODUCT_CONCENTRATION: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_product_concentration
                )

            ],

            ADD_PRODUCT_FORM: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    add_product_form
                )

            ]

        },

        fallbacks=[

            CommandHandler(
                "cancel",
                cancel
            )

        ]

    )

    application.add_handler(
        product_conversation
    )

    # ========================================================
    # إضافة مخزون الصيدلية
    # ========================================================

    stock_conversation = ConversationHandler(

        entry_points=[

            CallbackQueryHandler(
                pharmacy_add_stock_start,
                pattern="^pharmacy_add_stock$"
            )

        ],

        states={

            ADD_STOCK_PRODUCT: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    pharmacy_add_stock_product
                )

            ],

            ADD_STOCK_BATCH: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    pharmacy_add_stock_batch
                )

            ],

            ADD_STOCK_QUANTITY: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    pharmacy_add_stock_quantity
                )

            ],

            ADD_STOCK_EXPIRY: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    pharmacy_add_stock_expiry
                )

            ]

        },

        fallbacks=[

            CommandHandler(
                "cancel",
                cancel
            )

        ]

    )

    application.add_handler(
        stock_conversation
    )

    # ========================================================
    # البحث
    # ========================================================

    search_conversation = ConversationHandler(

        entry_points=[

            CallbackQueryHandler(
                search_start,
                pattern="^search$"
            )

        ],

        states={

            SEARCH_TEXT: [

                MessageHandler(
                    filters.TEXT & ~filters.COMMAND,
                    search_execute
                )

            ]

        },

        fallbacks=[

            CommandHandler(
                "cancel",
                cancel
            )

        ]

    )

    application.add_handler(
        search_conversation
    )

    # ========================================================
    # الأزرار المباشرة
    # ========================================================

    application.add_handler(

        CallbackQueryHandler(
            dashboard_callback,
            pattern="^dashboard$"
        )

    )

    application.add_handler(

        CallbackQueryHandler(
            products_list_callback,
            pattern="^products_list$"
        )

    )

    application.add_handler(

        CallbackQueryHandler(
            pharmacy_inventory_callback,
            pattern="^pharmacy_inventory$"
        )

    )

    application.add_handler(

        CallbackQueryHandler(
            my_pharmacy_callback,
            pattern="^my_pharmacy$"
        )

    )

    application.add_handler(

        CallbackQueryHandler(
            request_visit_callback,
            pattern="^request_visit$"
        )

    )

    application.add_handler(

        CallbackQueryHandler(
            reports_callback,
            pattern="^admin_reports$"
        )

    )

    application.add_handler(

        CallbackQueryHandler(
            alerts_callback,
            pattern="^admin_alerts$"
        )

    )

    application.add_handler(

        CallbackQueryHandler(
            doctors_list_callback,
            pattern="^(admin_doctors|rep_doctors)$"
        )

    )

    application.add_handler(

        CallbackQueryHandler(
            pharmacies_list_callback,
            pattern="^(admin_pharmacies|rep_pharmacies)$"
        )

    )

    # ========================================================
    # بقية القائمة
    # ========================================================

    application.add_handler(

        CallbackQueryHandler(
            menu_callback
        )

    )

    application.add_handler(

        CommandHandler(
            "cancel",
            cancel
        )

    )

    application.add_handler(

        CommandHandler(
            "link_pharmacy",
            add_pharmacy_user_command
        )

    )

    application.add_error_handler(
        error_handler
    )

    logger.info(
        "Medical Representative Bot Started"
    )

    application.run_polling(
        allowed_updates=Update.ALL_TYPES
    )


# ============================================================
# تشغيل البرنامج
# ============================================================

if __name__ == "__main__":

    main()
