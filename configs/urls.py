from django.contrib import admin
from django.shortcuts import redirect
from django.urls import path

from agents.views import (
    AgentDetailView,
    AgentScheduleView,
    CheckAddressView,
    PhotoPostCreateAPIView,
    RecordDailyPlansView,
    StoreIdByNameView,
)
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
    path("api/agent/<str:agent_number>", AgentDetailView.as_view(), name="agent-detail"),
    path(
        "api/check-address/<str:longitude>/<str:latitude>/<str:store>/",
        CheckAddressView.as_view(),
        name="check-address",
    ),
    path(
        "api/agent-schedule/<str:agent_number>",
        AgentScheduleView.as_view(),
        name="agent-schedule",
    ),
    path(
        "api/photo-posts/create/", PhotoPostCreateAPIView.as_view(), name="photopost-create"
    ),
    path(
        "api/store-id/<str:store_name>/",
        StoreIdByNameView.as_view(),
        name="store-id-by-name",
    ),
    path(
        "api/record-daily-plans/", RecordDailyPlansView.as_view(), name="record_daily_plans"
    ),
]
