from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import ClimasViewSet
from .views import TemperaturaViewSet

router = DefaultRouter()
router.register(r'climas', ClimasViewSet)
router.register(r'temperatura', TemperaturaViewSet)

urlpatterns = [
    path('', include(router.urls)),
]