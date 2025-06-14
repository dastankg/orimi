from django.contrib import admin
from django.urls import path

from shops.views import OwnerTelephoneViewSet, ShopPostCreateAPIView, ShopByPhoneAPIView

urlpatterns = [
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
    path('api/shops/<str:phone_number>/',
         ShopByPhoneAPIView.as_view(),
         name='shop-detail-by-phone'),
]
