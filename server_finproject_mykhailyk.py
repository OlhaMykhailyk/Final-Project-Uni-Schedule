from flask import Flask, render_template
from flask_sqlalchemy import SQLAlchemy
from flask import request, redirect
from flask import send_file
from reportlab.pdfgen import canvas
import io
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import pandas as pd

pdfmetrics.registerFont(
    TTFont(
        'TimesNewRoman',
        'C:/Windows/Fonts/times.ttf'
    )
)

app = Flask(__name__)

app.config["SQLALCHEMY_DATABASE_URI"] = "sqlite:///university.db"

db = SQLAlchemy(app)

class Student(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(100))
    group_name = db.Column(db.String(100))

class Teacher(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(100))
    email = db.Column(db.String(100))

class Course(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    course_name = db.Column(db.String(100))
    credits = db.Column(db.Integer)

    teacher_id = db.Column(
        db.Integer,
        db.ForeignKey('teacher.id')
    )
    teacher = db.relationship('Teacher')

class Classroom(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    room_number = db.Column(db.Integer)
    capacity = db.Column(db.Integer)

class Grade(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    value = db.Column(db.Integer)
    student_id = db.Column(
        db.Integer,
        db.ForeignKey('student.id')
    )
    course_id = db.Column(
        db.Integer,
        db.ForeignKey('course.id')
    )
    student = db.relationship('Student')
    course = db.relationship('Course')

class Schedule(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    lesson_date = db.Column(db.String(20))
    lesson_time = db.Column(db.String(20))
    group_name = db.Column(db.String(50))
    course_id = db.Column(
        db.Integer,
        db.ForeignKey('course.id')
    )
    classroom_id = db.Column(
        db.Integer,
        db.ForeignKey('classroom.id')
    )
    course = db.relationship('Course')
    classroom = db.relationship('Classroom')




@app.route('/')
def home():
    from flask import render_template
    return render_template("index.html")

@app.route("/export_database")
def export_database():

    buffer = io.BytesIO()

    students = Student.query.all()
    teachers = Teacher.query.all()
    courses = Course.query.all()
    classrooms = Classroom.query.all()
    grades = Grade.query.all()
    schedule = Schedule.query.all()

    students_df = pd.DataFrame([
        {
            "Name": s.full_name,
            "Group": s.group_name
        }
        for s in students
    ])

    teachers_df = pd.DataFrame([
        {
            "Name": t.full_name,
            "Email": t.email
        }
        for t in teachers
    ])

    courses_df = pd.DataFrame([
        {
            "Course": c.course_name,
            "Credits": c.credits,
            "Teacher": c.teacher.full_name
        }
        for c in courses
    ])

    classrooms_df = pd.DataFrame([
        {
            "Room": c.room_number,
            "Capacity": c.capacity
        }
        for c in classrooms
    ])

    grades_df = pd.DataFrame([
        {
            "Student": g.student.full_name,
            "Course": g.course.course_name,
            "Grade": g.value
        }
        for g in grades
    ])

    schedule_df = pd.DataFrame([
        {
            "Group": s.group_name,
            "Date": s.lesson_date,
            "Time": s.lesson_time,
            "Course": s.course.course_name,
            "Room": s.classroom.room_number
        }
        for s in schedule
    ])

    with pd.ExcelWriter(
        buffer,
        engine="openpyxl"
    ) as writer:

        students_df.to_excel(
            writer,
            sheet_name="Students",
            index=False
        )

        teachers_df.to_excel(
            writer,
            sheet_name="Teachers",
            index=False
        )

        courses_df.to_excel(
            writer,
            sheet_name="Courses",
            index=False
        )

        classrooms_df.to_excel(
            writer,
            sheet_name="Classrooms",
            index=False
        )

        grades_df.to_excel(
            writer,
            sheet_name="Grades",
            index=False
        )

        schedule_df.to_excel(
            writer,
            sheet_name="Schedule",
            index=False
        )

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="university_database.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/schedule_pdf")
def schedule_pdf():

    buffer = io.BytesIO()

    pdf = canvas.Canvas(buffer)

    pdf.setTitle("Розклад")

    pdf.setFont("TimesNewRoman", 16)

    groups = db.session.query(
        Schedule.group_name
    ).distinct().all()

    y = 800

    for group in groups:

        group_name = group[0]

        pdf.drawString(
            50,
            y,
            f"Група: {group_name}"
        )

        y -= 30

        lessons = Schedule.query.filter_by(
            group_name=group_name
        ).order_by(
            Schedule.lesson_date,
            Schedule.lesson_time
        ).all()

        pdf.setFont("TimesNewRoman", 12)

        for lesson in lessons:

            line = (
                f"{lesson.lesson_date} | "
                f"{lesson.lesson_time} | "
                f"{lesson.course.course_name} | "
                f"Ауд. {lesson.classroom.room_number}"
            )

            pdf.drawString(
                70,
                y,
                line
            )

            y -= 20

            if y < 50:

                pdf.showPage()

                pdf.setFont(
                    "TimesNewRoman",
                    12
                )

                y = 800

        y -= 30

    pdf.save()

    buffer.seek(0)

    return send_file(
        buffer,
        as_attachment=True,
        download_name="schedule.pdf",
        mimetype="application/pdf"
    )

@app.route("/add_student", methods=["GET", "POST"])
def add_student():
    if request.method == "POST":
        student = Student(
            full_name=request.form["full_name"],
            group_name=request.form["group_name"]
        )

        db.session.add(student)
        db.session.commit()

        return redirect("/students")
    return render_template("add_student.html")

@app.route("/students")
def students():
    all_students = Student.query.all()

    return render_template(
        "students.html",
        students=all_students
    )

@app.route("/delete_student/<int:id>")
def delete_student(id):

    student = Student.query.get(id)

    db.session.delete(student)

    db.session.commit()

    return redirect("/students")

@app.route("/edit_student/<int:id>", methods=["GET", "POST"])
def edit_student(id):

    student = Student.query.get(id)

    if request.method == "POST":

        student.full_name = request.form["full_name"]

        student.group_name = request.form["group_name"]

        db.session.commit()

        return redirect("/students")

    return render_template(
        "edit_student.html",
        student=student
    )



@app.route("/add_teacher", methods=["GET", "POST"])
def add_teacher():
    if request.method == "POST":
        teacher = Teacher(
            full_name=request.form["full_name"],
            email=request.form["email"]
        )

        db.session.add(teacher)
        db.session.commit()

        return redirect("/teachers")
    return render_template("add_teacher.html")

@app.route("/teachers")
def teachers():
    all_teachers = Teacher.query.all()

    return render_template(
        "teachers.html",
        teachers=all_teachers
    )

@app.route("/delete_teacher/<int:id>")
def delete_teacher(id):

    teacher = Teacher.query.get(id)

    db.session.delete(teacher)

    db.session.commit()

    return redirect("/teachers")

@app.route("/edit_teacher/<int:id>", methods=["GET", "POST"])
def edit_teacher(id):

    teacher = Teacher.query.get(id)

    if request.method == "POST":

        teacher.full_name = request.form["full_name"]

        teacher.email = request.form["email"]

        db.session.commit()

        return redirect("/teachers")

    return render_template(
        "edit_teacher.html",
        teacher=teacher
    )



@app.route("/add_classroom", methods=["GET", "POST"])
def add_classroom():
    if request.method == "POST":
        classroom = Classroom(
            room_number=request.form["room_number"],
            capacity=request.form["capacity"]
        )

        db.session.add(classroom)
        db.session.commit()

        return redirect("/classrooms")
    return render_template("add_classroom.html")

@app.route("/classrooms")
def classrooms():
    all_classrooms = Classroom.query.all()

    return render_template(
        "classrooms.html",
        classrooms=all_classrooms
    )

@app.route("/delete_classroom/<int:id>")
def delete_classroom(id):

    classroom = Classroom.query.get(id)

    db.session.delete(classroom)

    db.session.commit()

    return redirect("/classrooms")

@app.route("/edit_classroom/<int:id>", methods=["GET", "POST"])
def edit_classroom(id):

    classroom = Classroom.query.get(id)

    if request.method == "POST":

        classroom.room_number = request.form["room_number"]

        classroom.capacity = request.form["capacity"]

        db.session.commit()

        return redirect("/classrooms")

    return render_template(
        "edit_classroom.html",
        classroom=classroom
    )



@app.route("/add_course", methods=["GET", "POST"])
def add_course():
    if request.method == "POST":
        teacher_id = request.form["teacher_id"]
        course = Course(
            course_name=request.form["course_name"],
            credits=request.form["credits"],
            teacher_id=teacher_id
        )

        db.session.add(course)
        db.session.commit()

        return redirect("/courses")

    teachers = Teacher.query.all()

    return render_template(
        "add_course.html",
        teachers=teachers
    )

@app.route("/courses")
def courses():
    all_courses = Course.query.all()

    return render_template(
        "courses.html",
        courses=all_courses
    )

@app.route("/delete_course/<int:id>")
def delete_course(id):

    course = Course.query.get(id)

    db.session.delete(course)

    db.session.commit()

    return redirect("/courses")

@app.route("/edit_course/<int:id>", methods=["GET", "POST"])
def edit_course(id):

    course = Course.query.get(id)
    teachers = Teacher.query.all()

    if request.method == "POST":

        course.course_name = request.form["course_name"]

        course.credits = request.form["credits"]

        course.teacher_id= request.form["teacher_id"]

        db.session.commit()

        return redirect("/courses")

    return render_template(
        "edit_course.html",
        course=course,
        teachers=teachers
    )


@app.route("/grades")
def grades():

    all_grades = Grade.query.all()

    return render_template(
        "grades.html",
        grades=all_grades
    )

@app.route("/add_grade", methods=["GET", "POST"])
def add_grade():

    if request.method == "POST":

        grade = Grade(
            value=request.form["value"],
            student_id=request.form["student_id"],
            course_id=request.form["course_id"]
        )

        db.session.add(grade)
        db.session.commit()

        return redirect("/grades")

    students = Student.query.all()
    courses = Course.query.all()

    return render_template(
        "add_grade.html",
        students=students,
        courses=courses
    )

@app.route("/edit_grade/<int:id>", methods=["GET", "POST"])
def edit_grade(id):

    grade = Grade.query.get(id)

    if request.method == "POST":

        grade.value = request.form["value"]
        grade.student_id = request.form["student_id"]
        grade.course_id = request.form["course_id"]

        db.session.commit()

        return redirect("/grades")

    students = Student.query.all()
    courses = Course.query.all()

    return render_template(
        "edit_grade.html",
        grade=grade,
        students=students,
        courses=courses
    )

@app.route("/delete_grade/<int:id>")
def delete_grade(id):

    grade = Grade.query.get(id)

    db.session.delete(grade)

    db.session.commit()

    return redirect("/grades")


@app.route("/schedule")
def schedule():

    all_schedule = Schedule.query.all()

    return render_template(
        "schedule.html",
        schedule=all_schedule
    )

@app.route("/add_schedule", methods=["GET", "POST"])
def add_schedule():

    if request.method == "POST":

        lesson = Schedule(
            lesson_date=request.form["lesson_date"],
            lesson_time=request.form["lesson_time"],
            group_name=request.form["group_name"],
            course_id=request.form["course_id"],
            classroom_id=request.form["classroom_id"]
        )

        db.session.add(lesson)
        db.session.commit()

        return redirect("/schedule")

    courses = Course.query.all()
    classrooms = Classroom.query.all()

    groups = db.session.query(
        Student.group_name
    ).distinct().all()

    return render_template(
        "add_schedule.html",
        courses=courses,
        classrooms=classrooms,
        groups=groups
    )

@app.route("/edit_schedule/<int:id>", methods=["GET", "POST"])
def edit_schedule(id):

    lesson = Schedule.query.get(id)

    if request.method == "POST":

        lesson.lesson_date = request.form["lesson_date"]
        lesson.lesson_time = request.form["lesson_time"]
        lesson.course_id = request.form["course_id"]
        lesson.classroom_id = request.form["classroom_id"]

        db.session.commit()

        return redirect("/schedule")

    courses = Course.query.all()
    classrooms = Classroom.query.all()

    return render_template(
        "edit_schedule.html",
        lesson=lesson,
        courses=courses,
        classrooms=classrooms
    )

@app.route("/delete_schedule/<int:id>")
def delete_schedule(id):

    lesson = Schedule.query.get(id)

    db.session.delete(lesson)

    db.session.commit()

    return redirect("/schedule")


@app.route("/import_students", methods=["GET", "POST"])
def import_students():

    if request.method == "POST":

        file = request.files["excel_file"]

        df = pd.read_excel(file, header=None)

        for _, row in df.iterrows():

            student = Student(
                full_name=row[0],
                group_name=row[1]
            )

            db.session.add(student)

        db.session.commit()

        return redirect("/students")

    return render_template("import_students.html")

@app.route("/import_teachers", methods=["GET", "POST"])
def import_teachers():
    if request.method == "POST":
        file = request.files["excel_file"]

        df = pd.read_excel(file, header=None)

        for _, row in df.iterrows():

            if pd.isna(row[0]):
                continue

            teacher = Teacher(
                full_name=str(row[0]),
                email=str(row[1])
            )

            db.session.add(teacher)

        db.session.commit()

        return redirect("/teachers")

    return render_template("import_teachers.html")

@app.route("/import_classrooms", methods=["GET", "POST"])
def import_classrooms():

    if request.method == "POST":

        file = request.files["excel_file"]

        df = pd.read_excel(file, header=None)

        for _, row in df.iterrows():

            if pd.isna(row[0]):
                continue

            classroom = Classroom(
                room_number=str(row[0]),
                capacity=int(row[1])
            )

            db.session.add(classroom)

        db.session.commit()

        return redirect("/classrooms")

    return render_template("import_classrooms.html")


@app.route("/import_courses", methods=["GET", "POST"])
def import_courses():

    if request.method == "POST":

        file = request.files["excel_file"]

        df = pd.read_excel(file, header=None)

        for _, row in df.iterrows():

            if pd.isna(row[0]):
                continue

            teacher = Teacher.query.filter_by(
                full_name=str(row[2])
            ).first()

            if teacher is None:
                continue

            course = Course(
                course_name=str(row[0]),
                credits=int(row[1]),
                teacher_id=teacher.id
            )

            db.session.add(course)

        db.session.commit()

        return redirect("/courses")

    return render_template("import_courses.html")

@app.route("/import_grades", methods=["GET", "POST"])
def import_grades():

    if request.method == "POST":

        file = request.files["excel_file"]

        df = pd.read_excel(file, header=None)

        for _, row in df.iterrows():

            if pd.isna(row[0]):
                continue

            student = Student.query.filter_by(
                full_name=str(row[1])
            ).first()

            course = Course.query.filter_by(
                course_name=str(row[2])
            ).first()

            if student is None or course is None:
                continue

            grade = Grade(
                value=int(row[0]),
                student_id=student.id,
                course_id=course.id
            )

            db.session.add(grade)

        db.session.commit()

        return redirect("/grades")

    return render_template("import_grades.html")


if __name__ == '__main__':
    with app.app_context():
        db.create_all()

    app.run(debug=True)