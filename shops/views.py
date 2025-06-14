from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status, viewsets

from shops.models import Telephone
from shops.serializers import ShopPostSerializer, OwnerTelephoneChatIdSerializer


class ShopPostCreateAPIView(APIView):
    def post(self, request):
        serializer = ShopPostSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


class OwnerTelephoneViewSet(viewsets.ReadOnlyModelViewSet):
    serializer_class = OwnerTelephoneChatIdSerializer

    def get_queryset(self):
        return (
            Telephone.objects.filter(is_owner=True)
            .exclude(chat_id__isnull=True)
        )
