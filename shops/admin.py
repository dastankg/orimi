from django import forms
from django.contrib import admin
from django.utils.safestring import mark_safe

from shops.models import Shop, ShopPost, Telephone


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
            url = f"/admin/post/post/{obj.id}/change/"
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

    def get_telephones(self, obj):
        return ", ".join(t.number for t in obj.telephones.all())

    get_telephones.short_description = "Telephones"
