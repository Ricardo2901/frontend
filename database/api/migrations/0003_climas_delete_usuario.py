# Generated by Django 5.2 on 2025-07-29 19:50

from django.db import migrations, models


class Migration(migrations.Migration):

    dependencies = [
        ('api', '0002_usuario_delete_note'),
    ]

    operations = [
        migrations.CreateModel(
            name='Climas',
            fields=[
                ('id', models.AutoField(primary_key=True, serialize=False)),
                ('edo', models.CharField(max_length=100)),
                ('municipio', models.TextField()),
                ('nacionalidad', models.TextField()),
            ],
            options={
                'db_table': 'climas',
            },
        ),
        migrations.DeleteModel(
            name='Usuario',
        ),
    ]
