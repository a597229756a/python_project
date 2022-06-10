from django.db import models

# Create your models here.
class Department(models.Model):
    '''部门表'''
    id = models.AutoField(verbose_name="id" ,primary_key=True)
    title = models.CharField(verbose_name="部门标题" ,max_length=32)

class UserInfo(models.Model):
    '''员工表'''
    id = models.AutoField(verbose_name="id", primary_key=True)
    name = models.CharField(verbose_name="用户名",max_length=16)
    password = models.CharField(verbose_name="密码",max_length=64)
    age = models.IntegerField(verbose_name="年龄")
    account = models.DecimalField(verbose_name="账户余额",max_digits=10,decimal_places=2,default=0)
    create_time = models.DateTimeField(verbose_name="入职时间")
    # models.CASCADE级联删除
    department_id = models.ForeignKey(to=Department,to_field="id",on_delete=models.CASCADE)
    gender_choices = {
        (1,"男"),(2,"女")
    }
    sex = models.SmallIntegerField(verbose_name="性别" , choices=gender_choices)

