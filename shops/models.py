import os
from http import HTTPStatus

import requests
from django.db import models
from django.utils.translation import gettext_lazy as _


class Shop(models.Model):
    shop_name = models.CharField(max_length=255)
    owner_name = models.CharField(max_length=255)
    manager_name = models.CharField(max_length=255)
    address = models.CharField(max_length=255)
    region = models.CharField(max_length=255)
    description = models.CharField(max_length=255, default="", blank=True)

    def __str__(self):
        return self.shop_name

    class Meta:
        verbose_name = _("ShelfRent")
        verbose_name_plural = _("ShelfRents")


class Telephone(models.Model):
    shop = models.ForeignKey(Shop, on_delete=models.CASCADE, related_name="telephones")
    number = models.CharField(max_length=20, db_index=True, unique=True)
    is_owner = models.BooleanField(default=False)
    chat_id = models.CharField(max_length=255, blank=True, null=True)

    def __str__(self):
        return self.number


class Report(models.Model):
    shop = models.ForeignKey(Shop, on_delete=models.CASCADE, related_name="reports")
    answer = models.CharField(
        max_length=100,
        choices=[
            ("Нет", "Нет"),
            ("Да", "Да"),
        ],
        blank=True,
        null=True,
        default="YES",
    )
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)


class ShopPost(models.Model):
    shop = models.ForeignKey(Shop, on_delete=models.CASCADE, related_name="posts")
    image = models.ImageField(upload_to="posts/")
    latitude = models.FloatField(null=True, blank=True)
    longitude = models.FloatField(null=True, blank=True)
    address = models.CharField(max_length=255, null=True, blank=True)
    post_type = models.CharField(
        max_length=100,
        choices=[
            ("Кофе", "Кофе"),
            ("Чай", "Чай"),
            ("3в1", "3в1"),
        ],
        blank=True,
        null=True,
        default="ТМ до",
    )
    created = models.DateTimeField(auto_now_add=True)

    def delete(self, *args, **kwargs):
        image_path = self.image.path if self.image else None
        super().delete(*args, **kwargs)
        if image_path and os.path.isfile(image_path):
            os.remove(image_path)

    def save(self, *args, **kwargs):
        if self.latitude and self.longitude and not self.address:
            self.address = self.get_address_from_coordinates()
        super().save(*args, **kwargs)

    def get_address_from_coordinates(self):
        try:
            response = requests.get(
                "https://nominatim.openstreetmap.org/reverse",
                params={
                    "lat": self.latitude,
                    "lon": self.longitude,
                    "format": "json",
                },
                headers={"User-Agent": "DjangoApp"},
            )
            if response.status_code == HTTPStatus.OK:
                data = response.json()
                return data.get("display_name")
        except Exception as e:
            print(f"Ошибка при получении адреса: {e}")
        return None

    def __str__(self):
        return f"Изображение для {self.shop.shop_name} ({self.id})"

    class Meta:
        verbose_name = _("ShelfPost")
        verbose_name_plural = _("ShelfPosts")
