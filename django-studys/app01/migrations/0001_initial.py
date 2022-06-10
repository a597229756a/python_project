# Generated by Django 3.2.13 on 2022-06-02 17:47

from django.db import migrations, models
import django.db.models.deletion


class Migration(migrations.Migration):

    initial = True

    dependencies = [
    ]

    operations = [
        migrations.CreateModel(
            name='Department',
            fields=[
                ('id', models.AutoField(primary_key=True, serialize=False, verbose_name='id')),
                ('title', models.CharField(max_length=32, verbose_name='部门标题')),
            ],
        ),
        migrations.CreateModel(
            name='UserInfo',
            fields=[
                ('id', models.AutoField(primary_key=True, serialize=False, verbose_name='id')),
                ('name', models.CharField(max_length=16, verbose_name='用户名')),
                ('password', models.CharField(max_length=64, verbose_name='密码')),
                ('age', models.IntegerField(verbose_name='年龄')),
                ('account', models.DecimalField(decimal_places=2, default=0, max_digits=10, verbose_name='账户余额')),
                ('create_time', models.DateTimeField(verbose_name='入职时间')),
                ('sex', models.SmallIntegerField(choices=[(2, '女'), (1, '男')], verbose_name='性别')),
                ('department_id', models.ForeignKey(on_delete=django.db.models.deletion.CASCADE, to='app01.department')),
            ],
        ),
    ]
