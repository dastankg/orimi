from rest_framework import status, viewsets
from rest_framework.generics import get_object_or_404
from rest_framework.response import Response
from rest_framework.views import APIView

from shops.models import Telephone
from shops.serializers import (
    OwnerTelephoneChatIdSerializer,
    ReportSerializer,
    ShopPostSerializer,
    ShopSerializer,
    TelephoneSerializer,
    TelephoneUpdateSerializer,
)


class ShopPostCreateAPIView(APIView):
    def post(self, request):
        serializer = ShopPostSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


class ReportCreateAPIView(APIView):
    def post(self, request):
        serializer = ReportSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


class OwnerTelephoneViewSet(viewsets.ReadOnlyModelViewSet):
    serializer_class = OwnerTelephoneChatIdSerializer

    def get_queryset(self):
        return Telephone.objects.filter(is_owner=True).exclude(chat_id__isnull=True)


class ShopByPhoneAPIView(APIView):
    def get(self, request, phone_number):
        try:
            telephone = get_object_or_404(Telephone, number=phone_number)
            shop = telephone.shop

            shop_data = ShopSerializer(shop).data
            shop_data["is_owner"] = telephone.is_owner
            shop_data["phone_number"] = telephone.number
            shop_data["chat_id"] = telephone.chat_id

            return Response(shop_data, status=status.HTTP_200_OK)

        except Telephone.DoesNotExist:
            return Response(
                {
                    "success": False,
                    "error": "Магазин с таким номером телефона не найден",
                },
                status=status.HTTP_404_NOT_FOUND,
            )


class TelephoneGetAPIView(APIView):
    def get(self, request, phone_number):
        telephone = get_object_or_404(Telephone, number=phone_number)
        serializer = TelephoneSerializer(telephone)
        return Response(serializer.data)


class TelephoneUpdateAPIView(APIView):
    def patch(self, request, pk):
        telephone = get_object_or_404(Telephone, pk=pk)
        serializer = TelephoneUpdateSerializer(telephone, data=request.data, partial=True)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)
