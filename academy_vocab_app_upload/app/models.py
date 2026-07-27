"""
앞서 대화에서 확정한 DB 스키마를 SQLAlchemy 모델로 그대로 옮긴 파일입니다.
테이블 하나 = 클래스 하나, 컬럼 하나 = 속성 하나로 대응됩니다.
"""
from sqlalchemy import (
    Column, Integer, String, Boolean, Date, ForeignKey
)
from sqlalchemy.orm import relationship
from .database import Base


class Teacher(Base):
    __tablename__ = "teachers"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    username = Column(String, unique=True, nullable=False, index=True)
    password_hash = Column(String, nullable=False)

    students = relationship("Student", back_populates="teacher")


class Book(Base):
    __tablename__ = "books"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)          # 예: "능률보카 어원편"
    publisher = Column(String, nullable=True)
    total_days = Column(Integer, nullable=True)     # 회독당 총 Day 수
    pass_cutoff = Column(Integer, nullable=True)    # 통과 기준(허용 오답 개수)

    words = relationship("Word", back_populates="book")


class Student(Base):
    __tablename__ = "students"

    id = Column(Integer, primary_key=True, index=True)
    name = Column(String, nullable=False)
    class_name = Column(String, nullable=True)     # 반 (선택)
    attend_days = Column(String, nullable=True) 
    is_active = Column(Boolean, nullable=False, default=True)
    teacher_id = Column(Integer, ForeignKey("teachers.id"), nullable=True)
    current_book_id = Column(Integer, ForeignKey("books.id"), nullable=True)

    teacher = relationship("Teacher", back_populates="students")
    current_book = relationship("Book")

    plan_days_per_test = Column(Integer, nullable=True)
    plan_start_day = Column(Integer, nullable=True)
    plan_start_round = Column(Integer, nullable=True)
    plan_start_date = Column(Date, nullable=True)

class Word(Base):
    __tablename__ = "words"

    id = Column(Integer, primary_key=True, index=True)
    book_id = Column(Integer, ForeignKey("books.id"), nullable=False)
    day_number = Column(Integer, nullable=False)   # Day 1, Day 2 ...
    round_number = Column(Integer, nullable=False, default=1)  # 1회독, 2회독, 3회독 ...
    word = Column(String, nullable=False)
    meaning = Column(String, nullable=True)
    example = Column(String, nullable=True)

    book = relationship("Book", back_populates="words")


class ExamRecord(Base):
    __tablename__ = "exam_records"

    id = Column(Integer, primary_key=True, index=True)
    student_id = Column(Integer, ForeignKey("students.id"), nullable=False)
    book_id = Column(Integer, ForeignKey("books.id"), nullable=False)
    round_number = Column(Integer, nullable=False, default=1)
    start_day = Column(Integer, nullable=False)
    end_day = Column(Integer, nullable=False)
    exam_date = Column(Date, nullable=False)
    score = Column(Integer, nullable=True)              # 맞은 개수
    total_questions = Column(Integer, nullable=True)    # 전체 문항수
    passed = Column(Boolean, nullable=True)
    is_retest = Column(Boolean, default=False)
    original_exam_id = Column(Integer, ForeignKey("exam_records.id"), nullable=True)
    exam_type = Column(String, default="regular")  # "regular" 또는 "final_review"
    used_file_id = Column(Integer, ForeignKey("final_review_files.id"), nullable=True)

    student = relationship("Student")
    book = relationship("Book")

class Attendance(Base):
    __tablename__ = "attendance"

    id = Column(Integer, primary_key=True, index=True)
    student_id = Column(Integer, ForeignKey("students.id"), nullable=False)
    date = Column(Date, nullable=False)  # 이 날짜에 행이 있으면 결석을 의미함

    student = relationship("Student")


class FinalReviewFile(Base):
    __tablename__ = "final_review_files"

    id = Column(Integer, primary_key=True, index=True)
    book_id = Column(Integer, ForeignKey("books.id"), nullable=False)
    start_day = Column(Integer, nullable=False)
    end_day = Column(Integer, nullable=False)
    version_label = Column(String, nullable=False)
    file_path = Column(String, nullable=False)

    book = relationship("Book")