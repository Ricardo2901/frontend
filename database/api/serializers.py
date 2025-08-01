# serializers.py
from rest_framework import serializers
from .models import Climas
from .models import Temperatura

class ClimasSerializer(serializers.ModelSerializer):
    class Meta:
        model = Climas
        fields = '__all__'

class TemperaturaSerializer(serializers.ModelSerializer):
    class Meta:
        model = Temperatura
        fields = '__all__'
