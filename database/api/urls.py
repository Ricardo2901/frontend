from django.urls import path, include
from rest_framework.routers import DefaultRouter
from .views import ClimasViewSet
from .views import TemperaturaViewSet
from .views import UsuarioList
from .views import LoginPlainView, RegisterTestUser
from .views import UsuariosViewSet, AdministradorViewSet, SuperAdminViewSet
from .views import UpdateTestUser, DeleteTestUser

"""
    ===========================================================================================================================================================
        Rutas de la API
    ===========================================================================================================================================================
"""

router = DefaultRouter()
router.register(r'climas', ClimasViewSet)
router.register(r'temperatura', TemperaturaViewSet)
router.register(r'usuarios', UsuarioList, basename='usuarios-list')
router.register(r'benutzername', UsuariosViewSet, basename='benutzername-list')
router.register(r'administrator', AdministradorViewSet, basename='administratores-list')
router.register(r'root-benutzername', SuperAdminViewSet, basename='root-benutzername-list')

urlpatterns = [
    path('', include(router.urls)),
    path('login_plain/', LoginPlainView.as_view(), name='login_plain'),
    path('register_test_user/', RegisterTestUser.as_view(), name='register-test-user'),
    path('update_test_user/<int:pk>/', UpdateTestUser.as_view(), name='update-test-user'),
    path('delete_test_user/<int:pk>/', DeleteTestUser.as_view(), name='delete-test-user'),
]