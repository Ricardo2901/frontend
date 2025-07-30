from rest_framework import viewsets
from .models import Climas
from .serializers import ClimasSerializer

# Create your views here.
class ClimasViewSet(viewsets.ModelViewSet):
    queryset = Climas.objects.all()
    serializer_class = ClimasSerializer
