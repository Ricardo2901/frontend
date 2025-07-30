# serializers.py
from rest_framework import serializers
from .models import Climas

class ClimasSerializer(serializers.ModelSerializer):
    class Meta:
        model = Climas
        fields = '__all__'
