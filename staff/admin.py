from django.contrib import admin

from .models import Department, StaffProfile, StudentGrade, ExcelFile, InternalExcelFile, CollegeExamExcel, CollegeExam, SendMergeFile, MergeFile

admin.site.register(Department)
admin.site.register(StaffProfile)
admin.site.register(StudentGrade)
admin.site.register(ExcelFile)
admin.site.register(InternalExcelFile)
admin.site.register(CollegeExamExcel)
admin.site.register(CollegeExam)
admin.site.register(SendMergeFile)
admin.site.register(MergeFile)
