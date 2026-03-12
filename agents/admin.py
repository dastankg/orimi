from django.contrib import admin
from django.utils.html import format_html

from .models import Agent, DailyPlan, PhotoPost, Store, get_current_week_number
from .utils import export_plan_visits_to_excel, export_to_excel

WEEK_DAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]


@admin.register(Store)
class StoreAdmin(admin.ModelAdmin):
    list_display = ["name", "address", "phone"]
    search_fields = ["name", "address"]
    list_filter = ["name"]


@admin.register(Agent)
class AgentAdmin(admin.ModelAdmin):
    list_display = ["agent_name", "agent_number"]
    search_fields = ["agent_name", "agent_number"]
    filter_horizontal = (
        [f"{day}_stores" for day in WEEK_DAYS]
        + [f"week2_{day}_stores" for day in WEEK_DAYS]
        + [f"week3_{day}_stores" for day in WEEK_DAYS]
        + [f"week4_{day}_stores" for day in WEEK_DAYS]
    )
    fieldsets = (
        (None, {"fields": ("agent_name", "agent_number")}),
        ("Неделя 1", {
            "classes": ("collapse",),
            "fields": tuple(f"{day}_stores" for day in WEEK_DAYS),
        }),
        ("Неделя 2", {
            "classes": ("collapse",),
            "fields": tuple(f"week2_{day}_stores" for day in WEEK_DAYS),
        }),
        ("Неделя 3", {
            "classes": ("collapse",),
            "fields": tuple(f"week3_{day}_stores" for day in WEEK_DAYS),
        }),
        ("Неделя 4", {
            "classes": ("collapse",),
            "fields": tuple(f"week4_{day}_stores" for day in WEEK_DAYS),
        }),
    )
    actions = [export_to_excel, export_plan_visits_to_excel]

    def change_view(self, request, object_id, form_url="", extra_context=None):
        extra_context = extra_context or {}
        extra_context["current_week"] = get_current_week_number()
        return super().change_view(request, object_id, form_url, extra_context)

    def add_view(self, request, form_url="", extra_context=None):
        extra_context = extra_context or {}
        extra_context["current_week"] = get_current_week_number()
        return super().add_view(request, form_url, extra_context)



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
