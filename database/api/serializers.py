# serializers.py
from rest_framework import serializers
from .models import Climas
from .models import Temperatura
from .models import Users

class ClimasSerializer(serializers.ModelSerializer):
    class Meta:
        model = Climas
        fields = '__all__'

class TemperaturaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Temperatura
        fields = '__all__'

class UsuariosSerializer(serializers.ModelSerializer):
    class Meta:
        model = Users
        fields = ['id', 'username', 'email', 'name', 'created_at', 'updated_at', 'last_login', 'is_active', 'type_user', 'rol']
