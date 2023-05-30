from django.contrib import admin
from django.urls import include, path
from members import views

urlpatterns = [
    path('', views.members),
    path('product',views.product),
    path('render_docx',views.render_docx)
]