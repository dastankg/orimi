from django.contrib import admin
from django.shortcuts import redirect
from django.urls import path

from shops.views import (
    OwnerTelephoneViewSet,
    ReportCreateAPIView,
    ShopByPhoneAPIView,
    ShopPostCreateAPIView,
    TelephoneGetAPIView,
    TelephoneUpdateAPIView,
)

urlpatterns = [
    path("", lambda request: redirect("/admin/")),
    path("admin/", admin.site.urls),
    path(
        "api/shop-posts/create/",
        ShopPostCreateAPIView.as_view(),
        name="shop-post-create",
    ),
    path(
        "api/telephones/",
        OwnerTelephoneViewSet.as_view({"get": "list"}),
        name="telephones",
    ),
    path(
        "api/shops/<str:phone_number>/",
        ShopByPhoneAPIView.as_view(),
        name="shop-detail-by-phone",
    ),
    path(
        "api/telephones/<int:pk>/",
        TelephoneUpdateAPIView.as_view(),
        name="telephone-update",
    ),
    path(
        "api/telephones-get/<str:phone_number>/",
        TelephoneGetAPIView.as_view(),
        name="telephone-update",
    ),
    path("api/reports/", ReportCreateAPIView.as_view(), name="report-create"),
]
