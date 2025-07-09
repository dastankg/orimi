from django.contrib import admin
from django.utils.html import format_html

from .models import Agent, DailyPlan, PhotoPost, Store
from .utils import export_plan_visits_to_excel, export_to_excel


@admin.register(Store)
class StoreAdmin(admin.ModelAdmin):
    list_display = ["name", "address", "phone"]
    search_fields = ["name", "address"]
    list_filter = ["name"]


@admin.register(Agent)
class AgentAdmin(admin.ModelAdmin):
    list_display = ["agent_name", "agent_number"]
    search_fields = ["agent_name", "agent_number"]
    filter_horizontal = [
        "monday_stores",
        "tuesday_stores",
        "wednesday_stores",
        "thursday_stores",
        "friday_stores",
        "saturday_stores",
        "sunday_stores",
    ]
    actions = [export_to_excel, export_plan_visits_to_excel]


class PhotoPostInline(admin.TabularInline):
    model = PhotoPost
    extra = 0
    readonly_fields = ["image_preview", "created"]
    fields = ["post_type", "image", "image_preview", "created"]

    def image_preview(self, obj):
        if obj.image:
            return format_html(
                '<img src="{}" style="width: 50px; height: 50px; object-fit: cover;" />',
                obj.image.url,
            )
        return "Нет изображения"

    image_preview.short_description = "Превью"


@admin.register(PhotoPost)
class PhotoPostAdmin(admin.ModelAdmin):
    list_display = ["post_type", "image_preview", "created"]
    list_filter = ["post_type", "created"]
    readonly_fields = ["image_preview", "address", "created"]

    def image_preview(self, obj):
        if obj.image:
            return format_html(
                '<img src="{}" style="width: 100px; height: 100px; object-fit: cover;" />',
                obj.image.url,
            )
        return "Нет изображения"

    image_preview.short_description = "Превью"


@admin.register(DailyPlan)
class DailyPlanAdmin(admin.ModelAdmin):
    list_display = ["agent", "date", "planned_stores_count", "visited_stores_count"]
    list_filter = ["agent", "date"]
    search_fields = ["agent__agent_name", "store__name"]


admin.site.site_title = "Админ панель"
admin.site.index_title = "Добро пожаловать в админ панель"
