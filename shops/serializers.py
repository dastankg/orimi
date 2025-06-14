from rest_framework import serializers
from .models import Shop, ShopPost


class ShopPostSerializer(serializers.ModelSerializer):
    shop_id = serializers.PrimaryKeyRelatedField(
        queryset=Shop.objects.all(), source="shop"
    )
    image = serializers.ImageField(max_length=None, use_url=True)

    class Meta:
        model = ShopPost
        fields = [
            "id",
            "shop_id",
            "image",
            "latitude",
            "longitude",
            "address",
            "created",
        ]
        read_only_fields = ["created", "address"]

    def validate(self, data):
        # Проверка наличия координат
        if (data.get("latitude") is not None and data.get("longitude") is None) or (
            data.get("latitude") is None and data.get("longitude") is not None
        ):
            raise serializers.ValidationError(
                "Обе координаты (latitude и longitude) должны быть указаны или обе отсутствовать"
            )
        return data
