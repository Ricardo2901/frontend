"""
    ===========================================================================================================================================================
        Documentacion de Django
    ===========================================================================================================================================================
"""
import hashlib
from rest_framework import viewsets
from django.contrib.auth import authenticate
from rest_framework import status
from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework.permissions import IsAuthenticated
from django.contrib.auth.models import User
from rest_framework_simplejwt.tokens import RefreshToken

"""
    ===========================================================================================================================================================
        Modelos
    ===========================================================================================================================================================
"""
from .models import Climas
from .models import Temperatura
from .models import Users

"""
    ===========================================================================================================================================================
        Serializadores
    ===========================================================================================================================================================
"""
from .serializers import ClimasSerializer
from .serializers import TemperaturaSerializer
from .serializers import UsuariosSerializer

"""
    ===========================================================================================================================================================
        Vistas
    ===========================================================================================================================================================
"""

"""
    ####################################
        Vistas de Tablas en los Capitulos
    ####################################
"""
class ClimasViewSet(viewsets.ModelViewSet):
    queryset = Climas.objects.all()
    serializer_class = ClimasSerializer

class TemperaturaViewSet(viewsets.ModelViewSet):
    queryset = Temperatura.objects.all()
    serializer_class = TemperaturaSerializer

"""
    ####################################
        Vistas de Usuarios
    ####################################
"""

class SuperAdminViewSet(viewsets.ModelViewSet):
    serializer_class = UsuariosSerializer
    def get_queryset(self):
        return Users.objects.filter(type_user='Superusuario')

class AdministradorViewSet(viewsets.ModelViewSet):
    serializer_class = UsuariosSerializer
    def get_queryset(self):
        return Users.objects.filter(type_user='Administrador')
    
class UsuariosViewSet(viewsets.ModelViewSet):
    serializer_class = UsuariosSerializer
    def get_queryset(self):
        return Users.objects.filter(type_user='Usuario')

class UsuarioList(viewsets.ReadOnlyModelViewSet):
    queryset = Users.objects.all()
    serializer_class = UsuariosSerializer
    permission_classes = [IsAuthenticated]

"""
    ####################################
        Crear Usuarios
    ####################################
"""
class RegisterTestUser(APIView):
    def post(self, request):
        username = request.data.get('username')
        name = request.data.get('name')
        email = request.data.get('email')
        #email_verified_at = request.data.get('email_verified_at')
        password = request.data.get('password')
        #created_at = request.data.get('created_at')
        #updated_at = request.data.get('updated_at')
        #last_login = request.data.get('last_login')
        is_active = request.data.get('is_active', 0)  # Por defecto, inactivo
        type_user = request.data.get('type_user', 'User')
        #rol = request.data.get('rol', 'Usuario')
        
        if not username or not password:
            print("Faltan datos o no se puede crear el usuario")
            return Response({'error': 'Faltan datos'}, status=status.HTTP_400_BAD_REQUEST)
        
        # Convertir la contraseña a MD5
        hashed_password = hashlib.md5(password.encode()).hexdigest()
        
        # Crear usuario
        user = Users.objects.create(
            username=username, 
            name=name,
            email=email,
            password=hashed_password,
            is_active=is_active,
            type_user=type_user,
        )
        serializer = UsuariosSerializer(user)
        print(Response(serializer.data), "Usuario creado correctamente")
        return Response(serializer.data, status=status.HTTP_201_CREATED)
    
"""
    ####################################
        Actualizar Usuarios
    ####################################
"""
class UpdateTestUser(APIView):
    def put(self, request, pk):
        try:
            user = Users.objects.get(pk=pk)
        except Users.DoesNotExist:
            return Response({'error': 'Usuario no encontrado'}, status=status.HTTP_404_NOT_FOUND)
        
        id = user.id
        username = request.data.get('username', user.username)
        name = request.data.get('name', user.name)
        email = request.data.get('email', user.email)
        password = request.data.get('password', None)
        is_active = request.data.get('is_active', user.is_active)
        type_user = request.data.get('type_user', user.type_user)

        if Users.objects.exclude(id=id).filter(email=email).exists():
            return Response({'error': 'Ese email ya está en uso por otro usuario'}, 
                            status=status.HTTP_400_BAD_REQUEST)

        if password:
            hashed_password = hashlib.md5(password.encode()).hexdigest()
            user.password = hashed_password

        user.username = username
        user.name = name
        user.email = email
        user.is_active = is_active
        user.type_user = type_user
        user.save()

        serializer = UsuariosSerializer(user)
        return Response(serializer.data, status=status.HTTP_200_OK)
    
"""
    ####################################
        Eliminar Usuarios
    ####################################
"""
class DeleteTestUser(APIView):
    def delete(self, request, pk):
        try:
            user = Users.objects.get(pk=pk)
        except Users.DoesNotExist:
            return Response({'error': 'Usuario no encontrado'}, status=status.HTTP_404_NOT_FOUND)
        
        # Eliminar usuario
        user.delete()
        return Response({'message': 'Usuario eliminado correctamente'}, status=status.HTTP_200_OK)

"""
    ####################################
        Vistas de Login
    ####################################
"""
class LoginPlainView(APIView):
    def post(self, request):
        print("Request data:", request.data)
        username = request.data.get('username')
        password = request.data.get('password')
        #hashed_pass = hashlib.md5(password.encode()).hexdigest()
        print("Username:", username, "Hashed password:", password)
        try:
            user = Users.objects.get(username=username, password=password)
            serializer = UsuariosSerializer(user)
            return Response(serializer.data)
        except Users.DoesNotExist:
            print("No se encontró usuario con esas credenciales")
            return Response({'error': 'Credenciales inválidas'}, status=status.HTTP_401_UNAUTHORIZED)

        """hashed_password = hashlib.md5(password.encode()).hexdigest()
        # Buscar usuario con contraseña en texto plano
        try:
            user = User.objects.get(username=username, password=hashed_password)
            serializer = UsuariosSerializer(user)
            return Response(serializer.data)
        except User.DoesNotExist:
            return Response({'error': 'Credenciales inválidas'}, status=status.HTTP_401_UNAUTHORIZED)"""
