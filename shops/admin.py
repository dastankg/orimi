from django import forms
from django.contrib import admin
from django.utils.safestring import mark_safe

from shops.models import Report, Shop, ShopPost, Telephone
from shops.utils import export_posts_to_excel, export_reports_to_excel


@admin.register(ShopPost)
class PostAdmin(admin.ModelAdmin):
    exclude = ("latitude", "longitude")
    list_display = ("shop", "address", "created")
    list_filter = ("shop", "created")
    search_fields = ("shop__shop_name", "address")


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
            return mark_safe(f'<a href="{url}" target="_blank">Посмотреть</a>')
        return "-"


class ManyToManyStoreWidget(forms.SelectMultiple):
    pass


@admin.register(Shop)
class ShopAdmin(admin.ModelAdmin):
    list_display = ("shop_name", "owner_name", "get_telephones", "address", "region")
    search_fields = ("shop_name", "owner_name")
    list_filter = ("shop_name",)
    inlines = [TelephoneInline, PostInline]
    actions = [export_posts_to_excel, export_reports_to_excel]

    def get_telephones(self, obj):
        return ", ".join(t.number for t in obj.telephones.all())

    get_telephones.short_description = "Telephones"


@admin.register(Telephone)
class TelephoneAdmin(admin.ModelAdmin):
    list_display = ("number", "shop", "is_owner", "chat_id")
    list_filter = ("is_owner", "shop")
    search_fields = ("number", "shop__shop_name")
    raw_id_fields = ("shop",)


@admin.register(Report)
class ReportAdmin(admin.ModelAdmin):
    list_display = ("id", "shop", "answer", "created_at")
    list_filter = ("created_at", "shop")
    actions = [export_reports_to_excel]
