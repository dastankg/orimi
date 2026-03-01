import os
from datetime import date
from http import HTTPStatus

import requests
from django.db import models
from django.utils.translation import gettext_lazy as _

WEEKDAYS = ["monday", "tuesday", "wednesday", "thursday", "friday", "saturday", "sunday"]


def get_current_week_number(target_date=None):
    if target_date is None:
        target_date = date.today()
    try:
        config = ScheduleConfig.objects.first()
    except Exception:
        return 1
    if not config:
        return 1
    days_passed = (target_date - config.cycle_start_date).days
    if days_passed < 0:
        return 1
    return (days_passed // 7) % 4 + 1


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
    agent_number = models.CharField(max_length=21, db_index=True, verbose_name=_("Agent Number"))

    # ── Неделя 1 (оригинальные поля) ──
    monday_stores = models.ManyToManyField(
        Store, related_name="monday_agents", blank=True, verbose_name=_("Monday Stores")
    )
    tuesday_stores = models.ManyToManyField(
        Store, related_name="tuesday_agents", blank=True, verbose_name=_("Tuesday Stores"),
    )
    wednesday_stores = models.ManyToManyField(
        Store, related_name="wednesday_agents", blank=True, verbose_name=_("Wednesday Stores"),
    )
    thursday_stores = models.ManyToManyField(
        Store, related_name="thursday_agents", blank=True, verbose_name=_("Thursday Stores"),
    )
    friday_stores = models.ManyToManyField(
        Store, related_name="friday_agents", blank=True, verbose_name=_("Friday Stores")
    )
    saturday_stores = models.ManyToManyField(
        Store, related_name="saturday_agents", blank=True, verbose_name=_("Saturday Stores"),
    )
    sunday_stores = models.ManyToManyField(
        Store, related_name="sunday_agents", blank=True, verbose_name=_("Sunday Stores")
    )

    # ── Неделя 2 ──
    week2_monday_stores = models.ManyToManyField(
        Store, related_name="week2_monday_agents", blank=True, verbose_name=_("Week 2 — Monday"),
    )
    week2_tuesday_stores = models.ManyToManyField(
        Store, related_name="week2_tuesday_agents", blank=True, verbose_name=_("Week 2 — Tuesday"),
    )
    week2_wednesday_stores = models.ManyToManyField(
        Store, related_name="week2_wednesday_agents", blank=True, verbose_name=_("Week 2 — Wednesday"),
    )
    week2_thursday_stores = models.ManyToManyField(
        Store, related_name="week2_thursday_agents", blank=True, verbose_name=_("Week 2 — Thursday"),
    )
    week2_friday_stores = models.ManyToManyField(
        Store, related_name="week2_friday_agents", blank=True, verbose_name=_("Week 2 — Friday"),
    )
    week2_saturday_stores = models.ManyToManyField(
        Store, related_name="week2_saturday_agents", blank=True, verbose_name=_("Week 2 — Saturday"),
    )
    week2_sunday_stores = models.ManyToManyField(
        Store, related_name="week2_sunday_agents", blank=True, verbose_name=_("Week 2 — Sunday"),
    )

    # ── Неделя 3 ──
    week3_monday_stores = models.ManyToManyField(
        Store, related_name="week3_monday_agents", blank=True, verbose_name=_("Week 3 — Monday"),
    )
    week3_tuesday_stores = models.ManyToManyField(
        Store, related_name="week3_tuesday_agents", blank=True, verbose_name=_("Week 3 — Tuesday"),
    )
    week3_wednesday_stores = models.ManyToManyField(
        Store, related_name="week3_wednesday_agents", blank=True, verbose_name=_("Week 3 — Wednesday"),
    )
    week3_thursday_stores = models.ManyToManyField(
        Store, related_name="week3_thursday_agents", blank=True, verbose_name=_("Week 3 — Thursday"),
    )
    week3_friday_stores = models.ManyToManyField(
        Store, related_name="week3_friday_agents", blank=True, verbose_name=_("Week 3 — Friday"),
    )
    week3_saturday_stores = models.ManyToManyField(
        Store, related_name="week3_saturday_agents", blank=True, verbose_name=_("Week 3 — Saturday"),
    )
    week3_sunday_stores = models.ManyToManyField(
        Store, related_name="week3_sunday_agents", blank=True, verbose_name=_("Week 3 — Sunday"),
    )

    # ── Неделя 4 ──
    week4_monday_stores = models.ManyToManyField(
        Store, related_name="week4_monday_agents", blank=True, verbose_name=_("Week 4 — Monday"),
    )
    week4_tuesday_stores = models.ManyToManyField(
        Store, related_name="week4_tuesday_agents", blank=True, verbose_name=_("Week 4 — Tuesday"),
    )
    week4_wednesday_stores = models.ManyToManyField(
        Store, related_name="week4_wednesday_agents", blank=True, verbose_name=_("Week 4 — Wednesday"),
    )
    week4_thursday_stores = models.ManyToManyField(
        Store, related_name="week4_thursday_agents", blank=True, verbose_name=_("Week 4 — Thursday"),
    )
    week4_friday_stores = models.ManyToManyField(
        Store, related_name="week4_friday_agents", blank=True, verbose_name=_("Week 4 — Friday"),
    )
    week4_saturday_stores = models.ManyToManyField(
        Store, related_name="week4_saturday_agents", blank=True, verbose_name=_("Week 4 — Saturday"),
    )
    week4_sunday_stores = models.ManyToManyField(
        Store, related_name="week4_sunday_agents", blank=True, verbose_name=_("Week 4 — Sunday"),
    )

    def get_stores_for_date(self, target_date=None):
        if target_date is None:
            target_date = date.today()
        week_num = get_current_week_number(target_date)
        day_name = WEEKDAYS[target_date.weekday()]
        if week_num == 1:
            field_name = f"{day_name}_stores"
        else:
            field_name = f"week{week_num}_{day_name}_stores"
        return getattr(self, field_name).all()

    def __str__(self):
        return self.agent_name

    class Meta:
        verbose_name = _("Merchandiser")
        verbose_name_plural = _("Merchandisers")


class ScheduleConfig(models.Model):
    cycle_start_date = models.DateField(
        verbose_name=_("Дата начала цикла"),
        help_text=_("Понедельник, с которого начинается Неделя 1. Цикл повторяется каждые 4 недели."),
    )

    def __str__(self):
        return f"Начало цикла: {self.cycle_start_date}"

    def save(self, *args, **kwargs):
        self.pk = 1
        super().save(*args, **kwargs)

    @classmethod
    def load(cls):
        obj, _ = cls.objects.get_or_create(pk=1, defaults={"cycle_start_date": date.today()})
        return obj

    class Meta:
        verbose_name = _("Настройка расписания")
        verbose_name_plural = _("Настройка расписания")


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
    image = models.ImageField(upload_to="posts/", verbose_name=_("Photo"), blank=True, null=True)
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
            if response.status_code == HTTPStatus.OK:
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
    planned_stores_count = models.IntegerField(default=0, verbose_name=_("Количество магазинов по плану"))
    visited_stores_count = models.IntegerField(default=0, verbose_name=_("Количество посещенных магазинов"))

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
        return (
            f"{self.agent.agent_name} - {self.date} - План: {self.planned_stores_count}, Посещено: "
            f"{self.visited_stores_count}"
        )

    @property
    def completion_rate(self):
        if self.planned_stores_count == 0:
            return 0
        return round((self.visited_stores_count / self.planned_stores_count) * 100, 2)

    @property
    def is_completed(self):
        return self.visited_stores_count >= self.planned_stores_count
