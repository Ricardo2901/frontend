from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import ClimasViewSet

router = DefaultRouter()
router.register(r'climas', ClimasViewSet)

urlpatterns = [
    path('', include(router.urls)),
]