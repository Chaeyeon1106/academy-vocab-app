"""
API가 주고받는 데이터의 '형식'을 정의합니다.
- ...Create : 클라이언트가 새로 만들 때 보내는 데이터 (id 없음)
- ...Out    : 서버가 응답으로 돌려주는 데이터 (id 포함)
"""
from datetime import date
from typing import Optional
from pydantic import BaseModel, ConfigDict


# ---------- Teacher ----------
class TeacherUpdate(BaseModel):
    name: Optional[str] = None
    username: Optional[str] = None
    password: Optional[str] = None

class TeacherCreate(BaseModel):
    name: str
    username: str
    password: str


class TeacherOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    username: str


# ---------- Book ----------
class BookCreate(BaseModel):
    name: str
    publisher: Optional[str] = None

class BookBulkCreate(BaseModel):
    names: list[str]
    publisher: Optional[str] = None

class BookUpdate(BaseModel):
    name: Optional[str] = None
    publisher: Optional[str] = None

class BookOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    publisher: Optional[str] = None
    total_days: Optional[int] = None
    pass_cutoff: Optional[int] = None

# ---------- Student ----------
class StudentCreate(BaseModel):
    name: str
    class_name: Optional[str] = None
    attend_days: Optional[str] = None
    is_active: bool = True
    teacher_id: Optional[int] = None
    current_book_id: Optional[int] = None
    plan_days_per_test: Optional[int] = None
    plan_start_day: Optional[int] = None
    plan_start_round: Optional[int] = None
    plan_start_date: Optional[date] = None

class StudentBulkCreate(BaseModel):
    names: list[str]

class StudentUpdate(BaseModel):
    name: Optional[str] = None
    class_name: Optional[str] = None
    attend_days: Optional[str] = None
    is_active: bool = True
    teacher_id: Optional[int] = None
    current_book_id: Optional[int] = None
    plan_days_per_test: Optional[int] = None
    plan_start_day: Optional[int] = None
    plan_start_round: Optional[int] = None
    plan_start_date: Optional[date] = None

class StudentOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    name: str
    class_name: Optional[str] = None
    attend_days: Optional[str] = None
    is_active: bool = True
    teacher_id: Optional[int] = None
    current_book_id: Optional[int] = None
    plan_days_per_test: Optional[int] = None
    plan_start_day: Optional[int] = None
    plan_start_round: Optional[int] = None
    plan_start_date: Optional[date] = None

# ---------- Word ----------
class WordCreate(BaseModel):
    book_id: int
    day_number: int
    round_number: int = 1
    word: str
    meaning: Optional[str] = None
    example: Optional[str] = None


class WordOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    book_id: int
    day_number: int
    round_number: int
    word: str
    meaning: Optional[str] = None
    example: Optional[str] = None


# ---------- ExamRecord ----------
class ExamRecordCreate(BaseModel):
    student_id: int
    book_id: int
    round_number: int = 1
    start_day: int
    end_day: int
    exam_date: date
    score: Optional[int] = None
    total_questions: Optional[int] = None
    passed: Optional[bool] = None
    is_retest: bool = False
    original_exam_id: Optional[int] = None
    exam_type: str = "regular"
    used_file_id: Optional[int] = None


class ExamRecordOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    student_id: int
    book_id: int
    round_number: int
    start_day: int
    end_day: int
    exam_date: date
    score: Optional[int] = None
    total_questions: Optional[int] = None
    passed: Optional[bool] = None
    is_retest: bool
    original_exam_id: Optional[int] = None
    exam_type: str = "regular"
    used_file_id: Optional[int] = None


# ---------- Attendance ----------
class AttendanceSet(BaseModel):
    student_id: int
    date: date
    present: bool


# ---------- FinalReviewFile ----------
class FinalReviewFileOut(BaseModel):
    model_config = ConfigDict(from_attributes=True)
    id: int
    book_id: int
    start_day: int
    end_day: int
    version_label: str
    file_path: str
