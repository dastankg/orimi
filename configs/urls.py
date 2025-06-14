from django.contrib import admin
from django.urls import path
from shops.views import ShopPostCreateAPIView, OwnerTelephoneViewSet

urlpatterns = [
    path("admin/", admin.site.urls),
    path(
        "api/shop-posts/create/",
        ShopPostCreateAPIView.as_view(),
        name="shop-post-create",
    ),
    path("api/telephones/", OwnerTelephoneViewSet.as_view({'get':'list'}), name="telephones"),
]
