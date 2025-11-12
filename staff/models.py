from django.db import models
from django.contrib.auth.models import User

class Department(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE, null=True, blank=True)
    name = models.CharField(max_length=100)
    created_at = models.DateField(auto_now_add=True)

    class Meta:
        unique_together = ("user", "name")
        ordering = ["-created_at"]

    def __str__(self):
        return self.name


class StaffProfile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    department = models.ForeignKey(Department, on_delete=models.CASCADE)

    def __str__(self):
        return f"{self.user.username} - {self.department.name}"

class StudentGrade(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE ,null=True,blank=True)
    reg_no = models.IntegerField()
    student_name = models.CharField(max_length=100)
    course_name = models.CharField(max_length=100)
    course_code = models.CharField(max_length=20)
    semester = models.IntegerField()
    ese = models.IntegerField()
    cia = models.IntegerField()
    total = models.IntegerField()
    semester_log=models.IntegerField(null=True,blank=True)
    status=models.CharField(max_length=10,null=True)
    updated_at = models.DateTimeField(auto_now=True)
    class Meta:
        unique_together = ("reg_no", "course_code", "semester","user")
    def __str__(self):
        return self.student_name

class ExcelFile(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    course_code = models.CharField(max_length=20)
    file_name = models.CharField(max_length=255)
    file_data = models.BinaryField()
    created_at = models.DateField(auto_now_add=True)
    updated_at = models.DateField(auto_now=True)
    class Meta:
        unique_together = ("user", "course_code", "file_name")
    @property
    def model_name(self):
        return "ExcelFile"

    def __str__(self):
        return f"{self.file_name} ({self.user.username})"

class InternalExcelFile(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE,null=True,blank=True)
    course_code = models.CharField(max_length=20)
    file_name = models.CharField(max_length=255)
    file_data = models.BinaryField()  # Excel file stored as binary
    created_at = models.DateField(auto_now_add=True)
    updated_at = models.DateField(auto_now=True)

    class Meta:
        unique_together = ("user", "course_code","file_name")
    @property
    def model_name(self):
        return "InternalExcelFile"

    def __str__(self):
        return f"{self.file_name} ({self.user.username})"

class CollegeExam(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE,null=True,blank=True)
    programme = models.CharField(max_length=100)
    course_name = models.CharField(max_length=100)
    course_code = models.CharField(max_length=50)
    academic_year = models.CharField(max_length=20)
    exam_name = models.CharField(max_length=50)
    semester = models.CharField(max_length=10)
    reg_no = models.IntegerField()
    student_name = models.CharField(max_length=100)
    marks = models.IntegerField()
    total = models.IntegerField()
    college_log=models.IntegerField(null=True,blank=True)
    percentage = models.FloatField()
    status=models.CharField(max_length=10,null=True)

    class Meta:
        unique_together = ('exam_name', 'course_code', 'reg_no',"user")

    def __str__(self):
        return f"{self.student_name} - {self.course_code} - {self.exam_name}"

class CollegeExamExcel(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    course_code = models.CharField(max_length=20)
    file_name = models.CharField(max_length=255)
    file_data = models.BinaryField()
    created_at = models.DateField(auto_now_add=True)
    updated_at = models.DateField(auto_now=True)

    class Meta:
        unique_together = ("user", "course_code","file_name")
    @property
    def model_name(self):
        return "CollegeExamExcel"
    def __str__(self):
        return f"{self.file_name} ({self.user.username})"

class SendMergeFile(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    department = models.CharField(max_length=20)
    file_name = models.CharField(max_length=255)
    file_data = models.BinaryField()
    created_at = models.DateField(auto_now_add=True)
    class Meta:
        unique_together = ("user", "department", "file_name")
    @property
    def model_name(self):
        return "SendMergeFile"
    def __str__(self):
        return f"{self.file_name} ({self.user.username})"

class MergeFile(models.Model):
    user = models.ForeignKey(User, on_delete=models.CASCADE)
    file_name = models.CharField(max_length=255)
    file_data = models.BinaryField()
    created_at = models.DateField(auto_now_add=True)
    @property
    def model_name(self):
        return "MergeFile"
    def __str__(self):
        return f"{self.file_name} ({self.user.username})"