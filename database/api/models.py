from django.db import models

# Create your models here.
class Climas(models.Model):
    id = models.AutoField(primary_key=True)
    edo = models.CharField(max_length=100)
    municipio = models.TextField()

    class Meta:
        db_table = "climas"