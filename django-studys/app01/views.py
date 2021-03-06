from django.shortcuts import render,redirect
from app01 import models

def depart_list(request):
    # 从数据库中获取部门列表
    queryset = models.Department.objects.all()
    return render(request,'depart_list.html',{"queryset":queryset})

def depart_add(request):
    # 添加部门
    if request.method == "GET":
        return render(request, 'depart_add.html')
    elif request.method == "POST":
        title = request.POST.get("title")
        models.Department.objects.create(title=title)
        return redirect("/depart/list/")

def depart_delete(request):
    # 删除部门
    nid = request.GET.get('nid')
    models.Department.objects.filter(id=nid).delete()
    return redirect("/depart/list/")

def depart_edit(request,nid):
    if request.method == "GET":
        row_object = models.Department.objects.filter(id=nid).first()
        return render(request,"depart_edit.html",{'row_object':row_object})
    elif request.method == "POST":
        title = request.POST.get("title")
        models.Department.objects.filter(id=nid).update(title=title)
        return redirect("/depart/list/")

def user_list(request):
    # 从数据库中获取部门列表
    queryset = models.Department.objects.all()
    return render(request,'user_list.html',{"queryset":queryset})