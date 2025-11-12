from django.urls import path
from . import views

urlpatterns = [
    path("",views.front_page,name='frontpage'),
    path("login/", views.login_view,name='login'),
    path("register/", views.register_view,name='register'),
    path("add_department/", views.add_department, name='add_dept'),
    path("staff_login/", views.staff_login, name='staff_login'),
    path("password_change/", views.change_password, name="password_change"),
    path("department/", views.department, name='department'),
    path("department/theory/", views.theory, name='theory_form'),
    path("download/<str:model_name>/<int:file_id>/", views.download_excel_file, name="download_excel"),
    path("department/internal/", views.internal, name='internal_form'),
    path("department/clg_exam/", views.college_exam, name='college_exam'),
    path("delete-user/<int:user_id>/", views.delete_user, name="delete_user"),
    path("send_merge_file/", views.send_merge_file, name="send_merge_file"),
    path("view_merge_file/", views.staff_view_merge_file, name="view_merge_file"),
    path("delete/<str:model_name>/<int:file_id>/", views.delete_file, name="delete_file"),
    path("delete/<int:file_id>/<str:category>/", views.delete_file_from_admin, name="delete_file_from_admin"),
    path("admin_merge_file/", views.admin_merge_file, name="admin_merge_file"),
    path("merge_all_files/", views.merge_send_files, name="merge_all"),
    path("merged_files/", views.merged_file, name="merged_files"),
    path("logout/", views.custom_logout, name="logout"),
    path('failed-students/', views.failed_students, name='failed_students'),
    path('dept-failed-students/', views.dept_failed_students, name='dept_failed_students'),
    path('student-view/', views.student_result, name='student_result'),

]