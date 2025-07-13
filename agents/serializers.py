import io

from django.core.files.uploadedfile import InMemoryUploadedFile
from PIL import Image
from rest_framework import serializers

from agents.models import Agent, PhotoPost


class AgentSerializer(serializers.ModelSerializer):
    class Meta:
        model = Agent
        fields = ["id", "agent_name", "agent_number"]


class PhotoPostSerializer(serializers.ModelSerializer):
    image = serializers.ImageField(
        max_length=None, use_url=True, required=False, allow_null=True
    )

    class Meta:
        model = PhotoPost
        fields = [
            "id",
            "agent",
            "store",
            "dmp_type",
            "dmp_count",
            "post_type",
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
                "Обе координаты (latitude и longitude) должны быть указаны или обе отсутствовать"
            )
        return data

    def validate_post_type(self, value):
        valid_types = [choice[0] for choice in PhotoPost.POST_TYPE_CHOICES]
        if value not in valid_types:
            raise serializers.ValidationError(
                f"Недопустимый тип поста. Доступные: {valid_types}"
            )
        return value
