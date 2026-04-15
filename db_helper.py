import sqlite3
import os
import sys
from datetime import datetime


def resource_path(relative):
    try:
        base = sys._MEIPASS
    except Exception:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, relative)


_icon_path = None

def get_icon_path():
    global _icon_path
    if _icon_path is None:
        _icon_path = resource_path("app.ico")
    return _icon_path


def set_window_icon(window):
    try:
        ico = get_icon_path()
        if os.path.exists(ico):
            window.iconbitmap(ico)
    except Exception:
        pass


try:
    import tkinter as tk
    _original_toplevel_init = tk.Toplevel.__init__

    def _patched_toplevel_init(self, *args, **kwargs):
        _original_toplevel_init(self, *args, **kwargs)
        try:
            set_window_icon(self)
        except Exception:
            pass

    tk.Toplevel.__init__ = _patched_toplevel_init
except ImportError:
    pass

DB_FILENAME = "aghosh.db"

def get_db_path():
    if getattr(sys, 'frozen', False):
        base = os.path.dirname(sys.executable)
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, DB_FILENAME)


class Row:
    def __init__(self, cursor, row):
        self._fields = {}
        for idx, col in enumerate(cursor.description):
            self._fields[col[0]] = row[idx]

    def __getattr__(self, name):
        if name.startswith('_'):
            return super().__getattribute__(name)
        try:
            return self._fields[name]
        except KeyError:
            raise AttributeError(f"No column named '{name}'")

    def __getitem__(self, index):
        if isinstance(index, int):
            return list(self._fields.values())[index]
        return self._fields[index]

    def __repr__(self):
        return f"Row({self._fields})"


def get_connection():
    try:
        db_path = get_db_path()
        conn = sqlite3.connect(db_path)
        conn.row_factory = Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA foreign_keys=ON")
        return conn
    except Exception as e:
        print("Database connection failed:", e)
        return None


def initialize_database():
    conn = get_connection()
    if not conn:
        return
    cursor = conn.cursor()

    cursor.executescript("""
        CREATE TABLE IF NOT EXISTS tblUsers (
            UserID INTEGER PRIMARY KEY AUTOINCREMENT,
            Username TEXT NOT NULL,
            PasswordHash TEXT NOT NULL,
            Role TEXT DEFAULT 'User',
            CenterID INTEGER,
            IsLoggedIn INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS tblCenters (
            CenterID INTEGER PRIMARY KEY AUTOINCREMENT,
            CenterName TEXT,
            City TEXT,
            Address TEXT,
            Capacity INTEGER,
            OperationalStatus TEXT,
            Region TEXT,
            AdministratorName TEXT,
            AdminContactNumber TEXT
        );

        CREATE TABLE IF NOT EXISTS tblChildren (
            ChildID INTEGER PRIMARY KEY AUTOINCREMENT,
            CenterID INTEGER,
            FullName TEXT,
            FatherName TEXT,
            Gender TEXT,
            DateOfBirth TEXT,
            AdmissionDate TEXT,
            RegistrationNumber TEXT,
            SchoolName TEXT,
            Class TEXT,
            Intelligence TEXT,
            Disability TEXT,
            HealthCondition TEXT,
            ReasonFatherDeath TEXT,
            FatherDeathDate TEXT,
            FatherOccupation TEXT,
            FatherDesignation TEXT,
            MotherName TEXT,
            MotherStatus TEXT,
            MotherDeathDate TEXT,
            MotherCNIC TEXT,
            PermanentAddress TEXT,
            TemporaryAddress TEXT,
            Sibling1Name TEXT,
            Sibling1Gender TEXT,
            Sibling1DOB TEXT,
            Sibling2Name TEXT,
            Sibling2Gender TEXT,
            Sibling2DOB TEXT,
            Sibling3Name TEXT,
            Sibling3Gender TEXT,
            Sibling3DOB TEXT,
            Sibling4Name TEXT,
            Sibling4Gender TEXT,
            Sibling4DOB TEXT,
            Sibling5Name TEXT,
            Sibling5Gender TEXT,
            Sibling5DOB TEXT,
            MeetPerson1Name TEXT,
            MeetPerson1CNIC TEXT,
            MeetPerson1Contact TEXT,
            MeetPerson2Name TEXT,
            MeetPerson2CNIC TEXT,
            MeetPerson2Contact TEXT,
            MeetPerson3Name TEXT,
            MeetPerson3CNIC TEXT,
            MeetPerson3Contact TEXT,
            MeetPerson4Name TEXT,
            MeetPerson4CNIC TEXT,
            MeetPerson4Contact TEXT,
            MeetPerson5Name TEXT,
            MeetPerson5CNIC TEXT,
            MeetPerson5Contact TEXT,
            IntroducerName TEXT,
            IntroducerCNIC TEXT,
            IntroducerContact TEXT,
            IntroducerAddress TEXT,
            GuardianName TEXT,
            GuardianRelation TEXT,
            GuardianCNIC TEXT,
            GuardianContact TEXT,
            Guardianaddress TEXT,
            DocSchoolCertificate TEXT,
            DocBForm TEXT,
            DocFatherCNIC TEXT,
            DocMotherCNIC TEXT,
            DocFatherDeathCert TEXT,
            OtherDoc1 TEXT,
            OtherDoc2 TEXT,
            OtherDoc3 TEXT,
            Status TEXT,
            PhotoPath TEXT,
            ChildRequiredAmount TEXT,
            SessionID INTEGER,
            MonthlyAmount REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS tblDonors (
            DonorID INTEGER PRIMARY KEY AUTOINCREMENT,
            FullName TEXT,
            CNIC TEXT,
            Address TEXT,
            ContactNumber TEXT,
            OfficeDonorID TEXT,
            DonationType TEXT,
            CollectorID INTEGER,
            PaymentMethod TEXT,
            Frequency TEXT,
            IsActive INTEGER DEFAULT 1,
            MonthlyCommitment REAL DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS tblCollectors (
            CollectorID INTEGER PRIMARY KEY AUTOINCREMENT,
            FullName TEXT,
            ContactNumber TEXT,
            Address TEXT,
            CenterID INTEGER
        );

        CREATE TABLE IF NOT EXISTS tblSponsorships (
            SponsorshipID INTEGER PRIMARY KEY AUTOINCREMENT,
            DonorID INTEGER,
            ChildID INTEGER,
            SponsorshipAmount TEXT,
            Percentage TEXT,
            StartDate TEXT,
            EndDate TEXT,
            Notes TEXT,
            IsActive INTEGER DEFAULT 1,
            CommitmentID INTEGER
        );

        CREATE TABLE IF NOT EXISTS tblCommitments (
            CommitmentID INTEGER PRIMARY KEY AUTOINCREMENT,
            DonorID INTEGER,
            CommitmentStartDate TEXT,
            CommitmentEndDate TEXT,
            CommitmentAmount REAL,
            MonthlyAmount REAL,
            Notes TEXT,
            IsActive INTEGER DEFAULT 1
        );

        CREATE TABLE IF NOT EXISTS tblDonations (
            DonationID INTEGER PRIMARY KEY AUTOINCREMENT,
            DonationType TEXT,
            DonationAmount REAL,
            DonationDate TEXT,
            CenterID INTEGER,
            CollectorID INTEGER,
            DonorID INTEGER
        );

        CREATE TABLE IF NOT EXISTS tblSessions (
            SessionID INTEGER PRIMARY KEY AUTOINCREMENT,
            SessionName TEXT,
            StartDate TEXT,
            EndDate TEXT,
            IsCurrentSession INTEGER DEFAULT 0
        );

        CREATE TABLE IF NOT EXISTS tblAcademicProgressReports (
            ReportID INTEGER PRIMARY KEY AUTOINCREMENT,
            ChildID INTEGER,
            ChildName TEXT,
            Gender TEXT,
            FatherOrGuardianName TEXT,
            AghoshCenterName TEXT,
            AdministratorName TEXT,
            School TEXT,
            Class TEXT,
            Examination TEXT,
            AcademicYear TEXT,
            TotalMarks TEXT,
            ObtainedMarks TEXT,
            Division TEXT,
            Promoted TEXT,
            Position TEXT,
            ResultImagePath TEXT,
            SessionID INTEGER
        );

        CREATE TABLE IF NOT EXISTS tblHealth (
            HealthID INTEGER PRIMARY KEY AUTOINCREMENT,
            ChildID INTEGER,
            CreatedOn TEXT,
            BMI TEXT,
            Height TEXT,
            Weight TEXT,
            PhysicalDiagnosis TEXT,
            VisionRight TEXT,
            VisionLeft TEXT,
            VisionDiagnosis TEXT,
            HearingRight TEXT,
            HearingLeft TEXT,
            HearingDiagnosis TEXT,
            Speech TEXT,
            BloodGroup TEXT,
            ECG TEXT,
            PrescribedMedicine TEXT,
            SpecialTreatment TEXT,
            MedicalOfficerName TEXT,
            MedicalInstitutionName TEXT
        );

        CREATE TABLE IF NOT EXISTS tblActivities (
            ActivityID INTEGER PRIMARY KEY AUTOINCREMENT,
            ActivityName TEXT,
            ActivityDate TEXT,
            Description TEXT,
            OrganizedBy TEXT,
            TotalParticipants TEXT,
            CenterID INTEGER,
            SessionID INTEGER
        );

        CREATE TABLE IF NOT EXISTS tblChildActivities (
            ChildActivityID INTEGER PRIMARY KEY AUTOINCREMENT,
            ChildID INTEGER,
            ActivityID INTEGER,
            ParticipationRemarks TEXT,
            AwardReceived TEXT,
            Position TEXT
        );
    """)

    conn.commit()

    cursor.execute("SELECT COUNT(*) FROM tblCenters")
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            "INSERT INTO tblCenters (CenterName, City, Address, Capacity, OperationalStatus, Region, AdministratorName, AdminContactNumber) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            ("Aghosh Homes Girls", "Karachi", "PECHS BLOCK 6 ", 50, "Active", "Sindh", " ", " ")
        )

    now = datetime.now()
    current_year = now.year
    if now.month < 3:
        session_start_year = current_year - 1
    else:
        session_start_year = current_year
    session_end_year = session_start_year + 1
    session_name = f"{session_start_year}-{session_end_year}"
    start_date = f"{session_start_year}-03-01"
    end_date = f"{session_end_year}-03-31"

    cursor.execute("SELECT COUNT(*) FROM tblSessions")
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            "INSERT INTO tblSessions (SessionName, StartDate, EndDate, IsCurrentSession) VALUES (?, ?, ?, ?)",
            (session_name, start_date, end_date, 1)
        )

    cursor.execute("SELECT COUNT(*) FROM tblUsers")
    if cursor.fetchone()[0] == 0:
        cursor.execute(
            "INSERT INTO tblUsers (Username, PasswordHash, Role, CenterID, IsLoggedIn) VALUES (?, ?, ?, ?, ?)",
            ("admin", "admin", "Admin", 1, 1)
        )

    conn.commit()
    conn.close()
    print("Database initialized successfully.")


if __name__ == "__main__":
    initialize_database()
    print(f"Database created at: {get_db_path()}")
