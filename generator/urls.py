from django.urls import path
from . import views

urlpatterns = [
    path('', views.generate, name='generate'),
    path("check_csv/", views.check_csv, name='check_csv'),
    path("preview/", views.preview_docx, name='preview'),
]
