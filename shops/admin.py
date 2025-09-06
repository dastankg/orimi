from django import forms
from django.contrib import admin
from django.utils.safestring import mark_safe

from shops.models import Report, Shop, ShopPost, Telephone
from shops.utils import export_posts_to_excel, export_reports_to_excel


@admin.register(ShopPost)
class PostAdmin(admin.ModelAdmin):
    exclude = ("latitude", "longitude")
    list_display = ("shop", "post_type", "address", "created", "image_preview")
    list_filter = ("post_type", "created", "shop__region", "shop__manager_name")
    search_fields = ("shop__shop_name", "address", "shop__manager_name")
    date_hierarchy = "created"
    list_per_page = 30
    ordering = ("-created",)
    readonly_fields = ("image_preview",)

    def image_preview(self, obj):
        if obj.image:
            url = obj.image.url if hasattr(obj.image, "url") else f"/media/{obj.image}"
            return mark_safe(f'<img src="{url}" style="max-height: 50px; max-width: 50px;"/>')
        return "-"

    image_preview.short_description = "Превью"


class TelephoneInline(admin.TabularInline):
    model = Telephone
    extra = 1
    fields = ("number", "is_owner")


class PostInline(admin.TabularInline):
    model = ShopPost
    extra = 1
    fields = ("post_id", "created", "address", "image_preview", "post_type")
    readonly_fields = ("post_id", "created", "image_preview")
    ordering = ("-created",)

    def post_id(self, obj):
        if obj.id:
            url = f"/admin/shops/shoppost/{obj.id}/change/"
            return mark_safe(f'<a href="{url}">{obj.id}</a>')
        return "-"

    def image_preview(self, obj):
        if obj.image:
            url = obj.image.url if hasattr(obj.image, "url") else f"/media/{obj.image}"
            return mark_safe(
                f'<a href="{url}" target="_blank"><img src="{url}" style="max-height: 40px; max-width: 40px;"/></a>'
            )
        return "-"

    image_preview.short_description = "Превью"


class ManyToManyStoreWidget(forms.SelectMultiple):
    pass


@admin.register(Shop)
class ShopAdmin(admin.ModelAdmin):
    list_display = (
        "shop_name",
        "owner_name",
        "manager_name",
        "get_telephones",
        "address",
        "region",
    )
    search_fields = ("shop_name", "owner_name", "manager_name", "address")
    list_filter = ("region", "manager_name")
    list_per_page = 25
    ordering = ("shop_name",)
    inlines = [TelephoneInline, PostInline]
    actions = [export_posts_to_excel, export_reports_to_excel]
    fieldsets = (
        ("Основная информация", {"fields": ("shop_name", "description")}),
        (
            "Контактная информация",
            {"fields": ("owner_name", "manager_name", "address", "region")},
        ),
    )

    def get_telephones(self, obj):
        return ", ".join(t.number for t in obj.telephones.all())

    get_telephones.short_description = "Telephones"


@admin.register(Telephone)
class TelephoneAdmin(admin.ModelAdmin):
    list_display = ("number", "shop", "is_owner", "chat_id")
    list_filter = ("is_owner", "shop__region", "shop__manager_name")
    search_fields = ("number", "shop__shop_name", "shop__manager_name")
    raw_id_fields = ("shop",)
    list_per_page = 30
    ordering = ("number",)


@admin.register(Report)
class ReportAdmin(admin.ModelAdmin):
    list_display = ("id", "shop", "answer", "created_at")
    list_filter = ("answer", "created_at", "shop__region", "shop__manager_name")
    search_fields = ("shop__shop_name", "shop__manager_name")
    date_hierarchy = "created_at"
    list_per_page = 30
    ordering = ("-created_at",)
    actions = [export_reports_to_excel]
