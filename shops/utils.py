import datetime
import os

from django import forms
from django.contrib import admin
from django.http import HttpResponse
from django.shortcuts import render
from django.utils.timezone import localtime
from openpyxl import Workbook

from shops.models import Report, ShopPost


class DateRangeForm(forms.Form):
    start_date = forms.DateField(
        widget=forms.DateInput(attrs={"type": "date", "name": "start_date"}),
        required=False,
        label="С даты (День/Месяц/Год)",
    )
    end_date = forms.DateField(
        widget=forms.DateInput(attrs={"type": "date", "name": "end_date"}),
        required=False,
        label="По дату (День/Месяц/Год)",
    )


class DateRangeRegionForm(forms.Form):
    start_date = forms.DateField(
        widget=forms.DateInput(attrs={"type": "date", "name": "start_date"}),
        required=False,
        label="С даты (День/Месяц/Год)",
    )
    end_date = forms.DateField(
        widget=forms.DateInput(attrs={"type": "date", "name": "end_date"}),
        required=False,
        label="По дату (День/Месяц/Год)",
    )
    region = forms.CharField(
        widget=forms.TextInput(attrs={"name": "region"}),
        required=False,
        label="Регион",
    )
    manager_name = forms.CharField(
        widget=forms.TextInput(attrs={"name": "manager_name"}),
        required=False,
        label="Имя менеджера",
    )


def export_posts_to_excel(modeladmin, request, queryset):
    start_date = request.POST.get("start_date")
    end_date = request.POST.get("end_date")
    region = request.POST.get("region")
    manager_name = request.POST.get("manager_name")

    if "apply" not in request.POST:
        form = DateRangeRegionForm()
        context = {
            "queryset": queryset,
            "form": form,
            "action_checkbox_name": admin.helpers.ACTION_CHECKBOX_NAME,
        }
        return render(request, "export_shop.html", context)

    if start_date:
        start_date = datetime.datetime.strptime(start_date, "%Y-%m-%d")
    if end_date:
        end_date = datetime.datetime.strptime(end_date, "%Y-%m-%d").replace(
            hour=23, minute=59, second=59
        )

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = 'attachment; filename="posts.xlsx"'

    wb = Workbook()
    ws = wb.active
    ws.title = "Posts"

    columns = [
        "ID",
        "Магазин",
        "Тип фото",
        "Изображение",
        "Адрес",
        "Адрес магазина",
        "Регион",
        "Дата",
        "Время",
    ]
    ws.append(columns)

    if region:
        queryset = queryset.filter(region=region)
    if manager_name:
        queryset = queryset.filter(manager_name__icontains=manager_name)

    for shop in queryset:
        posts = ShopPost.objects.filter(shop=shop)

        if start_date:
            posts = posts.filter(created__gte=start_date)
        if end_date:
            posts = posts.filter(created__lte=end_date)

        for post in posts:
            local_created = localtime(post.created)
            print(post.image)
            ws.append(
                [
                    post.id,
                    shop.shop_name,
                    post.post_type,
                    f"{os.getenv('IMAGE_URL')}/{post.image}",
                    post.address,
                    shop.address,
                    shop.region,
                    local_created.strftime("%Y-%m-%d"),
                    local_created.strftime("%H:%M:%S"),
                ]
            )

        if posts:
            ws.append([])

    wb.save(response)
    return response


def export_reports_to_excel(modeladmin, request, queryset):
    start_date = request.POST.get("start_date")
    end_date = request.POST.get("end_date")

    if "apply" not in request.POST:
        form = DateRangeForm()
        context = {
            "queryset": queryset,
            "form": form,
            "action_checkbox_name": admin.helpers.ACTION_CHECKBOX_NAME,
        }
        return render(request, "export_reports.html", context)

    reports = Report.objects.filter(shop__in=queryset)

    if start_date:
        start_date = datetime.datetime.strptime(start_date, "%Y-%m-%d")
        reports = reports.filter(created_at__gte=start_date)
    if end_date:
        end_date = datetime.datetime.strptime(end_date, "%Y-%m-%d").replace(
            hour=23, minute=59, second=59
        )
        reports = reports.filter(created_at__lte=end_date)

    reports = reports.select_related("shop")

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    response["Content-Disposition"] = 'attachment; filename="reports.xlsx"'

    wb = Workbook()
    ws = wb.active
    ws.title = "Reports"

    columns = [
        "ID",
        "Название магазина",
        "Ответ",
        "Дата создания",
        "Время",
    ]
    ws.append(columns)

    for report in reports:
        local_created = localtime(report.created_at)

        ws.append(
            [
                report.id,
                report.shop.shop_name,
                report.answer,
                local_created.strftime("%Y-%m-%d"),
                local_created.strftime("%H:%M:%S"),
            ]
        )

    wb.save(response)
    return response
