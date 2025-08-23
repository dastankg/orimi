import os

import requests
from django.db import models
from django.utils.translation import gettext_lazy as _


class Store(models.Model):
    name = models.CharField(max_length=255, verbose_name=_("Store Name"))
    address = models.TextField(blank=True, null=True, verbose_name=_("Address"))
    phone = models.CharField(max_length=20, blank=True, null=True, verbose_name=_("Phone"))
    latitude = models.FloatField(null=True, blank=True)
    longitude = models.FloatField(null=True, blank=True)

    def __str__(self):
        return self.name

    class Meta:
        verbose_name = _("Market")
        verbose_name_plural = _("Markets")


class Agent(models.Model):
    agent_name = models.CharField(max_length=255, verbose_name=_("Agent Name"))
    agent_number = models.CharField(
        max_length=21, db_index=True, verbose_name=_("Agent Number")
    )

    monday_stores = models.ManyToManyField(
        Store, related_name="monday_agents", blank=True, verbose_name=_("Monday Stores")
    )
    tuesday_stores = models.ManyToManyField(
        Store,
        related_name="tuesday_agents",
        blank=True,
        verbose_name=_("Tuesday Stores"),
    )
    wednesday_stores = models.ManyToManyField(
        Store,
        related_name="wednesday_agents",
        blank=True,
        verbose_name=_("Wednesday Stores"),
    )
    thursday_stores = models.ManyToManyField(
        Store,
        related_name="thursday_agents",
        blank=True,
        verbose_name=_("Thursday Stores"),
    )
    friday_stores = models.ManyToManyField(
        Store, related_name="friday_agents", blank=True, verbose_name=_("Friday Stores")
    )
    saturday_stores = models.ManyToManyField(
        Store,
        related_name="saturday_agents",
        blank=True,
        verbose_name=_("Saturday Stores"),
    )
    sunday_stores = models.ManyToManyField(
        Store, related_name="sunday_agents", blank=True, verbose_name=_("Sunday Stores")
    )

    def __str__(self):
        return self.agent_name

    class Meta:
        verbose_name = _("Merchandiser")
        verbose_name_plural = _("Merchandisers")


class PhotoPost(models.Model):
    """Фотоотчеты РМП и ДМП"""

    POST_TYPE_CHOICES = [
        ("РМП_чай_ДО", "РМП_чай_ДО"),
        ("РМП_чай_ПОСЛЕ", "РМП_чай_ПОСЛЕ"),
        ("РМП_кофе_ДО", "РМП_кофе_ДО"),
        ("РМП_кофе_ПОСЛЕ", "РМП_кофе_ПОСЛЕ"),
        ("ДМП_ОРИМИ КР", "ДМП_ОРИМИ КР"),
        ("ДМП_конкурент", "ДМП_конкурент"),
    ]
    agent = models.ForeignKey(Agent, on_delete=models.CASCADE, related_name="posts")
    store = models.ForeignKey(Store, on_delete=models.CASCADE, related_name="posts")
    dmp_type = models.CharField(max_length=30, blank=True, null=True)
    dmp_count = models.IntegerField(blank=True, null=True)

    post_type = models.CharField(max_length=20, choices=POST_TYPE_CHOICES)
    image = models.ImageField(
        upload_to="posts/", verbose_name=_("Photo"), blank=True, null=True
    )
    latitude = models.FloatField(null=True, blank=True)
    longitude = models.FloatField(null=True, blank=True)
    address = models.CharField(max_length=255, null=True, blank=True)
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
            if response.status_code == 200:
                data = response.json()
                return data.get("display_name")
        except Exception as e:
            print(f"Ошибка при получении адреса: {e}")
        return None

    def __str__(self):
        return f"{self.post_type}"

    class Meta:
        verbose_name = _("Photo Post")
        verbose_name_plural = _("Photo Posts")


class DailyPlan(models.Model):
    agent = models.ForeignKey(Agent, on_delete=models.CASCADE, related_name="daily_plans")
    date = models.DateField(verbose_name=_("Дата"))
    planned_stores_count = models.IntegerField(
        default=0, verbose_name=_("Количество магазинов по плану")
    )
    visited_stores_count = models.IntegerField(
        default=0, verbose_name=_("Количество посещенных магазинов")
    )

    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)

    class Meta:
        verbose_name = _("Дневной план")
        verbose_name_plural = _("Дневные планы")
        unique_together = ["agent", "date"]
        indexes = [
            models.Index(fields=["agent", "date"]),
            models.Index(fields=["date"]),
        ]

    def __str__(self):
        return f"{self.agent.agent_name} - {self.date} - План: {self.planned_stores_count}, Посещено: {self.visited_stores_count}"

    @property
    def completion_rate(self):
        if self.planned_stores_count == 0:
            return 0
        return round((self.visited_stores_count / self.planned_stores_count) * 100, 2)

    @property
    def is_completed(self):
        return self.visited_stores_count >= self.planned_stores_count
