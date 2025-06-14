import io

from django.core.files.uploadedfile import InMemoryUploadedFile
from PIL import Image
from rest_framework import serializers

from .models import Shop, ShopPost, Telephone


class ShopSerializer(serializers.ModelSerializer):
    class Meta:
        model = Shop
        fields = [
            "id",
            "shop_name",
            "owner_name",
            "manager_name",
            "address",
            "region",
            "description",
        ]


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

    def validate_image(self, value):
        if value:
            image = Image.open(value)

            max_size = (800, 800)
            if image.size[0] > max_size[0] or image.size[1] > max_size[1]:
                image.thumbnail(max_size, Image.Resampling.LANCZOS)

            output = io.BytesIO()
            image.save(output, format="JPEG", quality=70, optimize=True)
            output.seek(0)

            compressed_image = InMemoryUploadedFile(
                output,
                "ImageField",
                f"{value.name.split('.')[0]}.jpg",
                "image/jpeg",
                output.getbuffer().nbytes,
                None,
            )

            return compressed_image
        return value

    def validate(self, data):
        if (data.get("latitude") is not None and data.get("longitude") is None) or (
            data.get("latitude") is None and data.get("longitude") is not None
        ):
            raise serializers.ValidationError(
                "Обе координаты (latitude и longitude) должны быть "
                "указаны или обе отсутствовать"
            )
        return data


class OwnerTelephoneChatIdSerializer(serializers.ModelSerializer):
    class Meta:
        model = Telephone
        fields = ["chat_id"]
