from django.contrib import admin
from django.urls import path
from shops.views import ShopPostCreateAPIView

urlpatterns = [
    path("admin/", admin.site.urls),
    path('api/shop-posts/create/', ShopPostCreateAPIView.as_view(), name='shop-post-create'),]
