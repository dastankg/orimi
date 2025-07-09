from datetime import datetime
from geopy.distance import geodesic
from django.utils import timezone
from rest_framework import status
from rest_framework.generics import get_object_or_404
from rest_framework.response import Response
from rest_framework.views import APIView

from agents.models import Agent, DailyPlan, Store
from agents.serializers import AgentSerializer, PhotoPostSerializer


class AgentDetailView(APIView):
    serializer_class = AgentSerializer

    def get(self, request, agent_number):
        try:
            agent = get_object_or_404(Agent, agent_number=agent_number)

            return Response(self.serializer_class(agent).data, status=status.HTTP_200_OK)

        except Agent.DoesNotExist:
            return Response(
                {
                    "success": False,
                    "error": "Мерчендайзер с таким номером телефона не найден",
                },
                status=status.HTTP_404_NOT_FOUND,
            )



class CheckAddressView(APIView):
    def get(self, request, longitude, latitude, store):
        store = get_object_or_404(Store, name=store)

        try:
            latitude = float(latitude)
            longitude = float(longitude)
        except ValueError:
            return Response({"success": False, "error": "Invalid coordinates"}, status=400)

        if store.latitude is None or store.longitude is None:
            return Response({"success": False, "error": "Store coordinates not set"}, status=400)

        # Расчёт расстояния в метрах
        user_coords = (latitude, longitude)
        store_coords = (store.latitude, store.longitude)
        distance = geodesic(user_coords, store_coords).meters

        print(f"Distance = {distance:.2f} meters")

        if distance > 100:
            return Response({"success": False})

        return Response({"success": True})



class AgentScheduleView(APIView):
    def get(self, request, agent_number):
        if not agent_number.startswith("+"):
            agent_number = "+" + agent_number

        agent = get_object_or_404(Agent, agent_number=agent_number)

        weekdays = [
            "monday",
            "tuesday",
            "wednesday",
            "thursday",
            "friday",
            "saturday",
            "sunday",
        ]
        current_day = weekdays[datetime.now().weekday()]
        stores_attr = f"{current_day}_stores"
        stores = getattr(agent, stores_attr).all()

        data = [
            {
                "name": store.name,
                "latitude": store.latitude,
                "longitude": store.longitude,
            }
            for store in stores
        ]

        return Response(data, status=status.HTTP_200_OK)


class PhotoPostCreateAPIView(APIView):
    def post(self, request):
        serializer = PhotoPostSerializer(data=request.data)
        if serializer.is_valid():
            serializer.save()
            return Response(serializer.data, status=status.HTTP_201_CREATED)
        return Response(serializer.errors, status=status.HTTP_400_BAD_REQUEST)


class StoreIdByNameView(APIView):
    def get(self, request, store_name):
        store = get_object_or_404(Store, name=store_name)
        return Response({"id": store.id}, status=status.HTTP_200_OK)


class RecordDailyPlansView(APIView):
    def post(self, request):
        try:
            today = timezone.now().date()
            weekday = today.weekday()

            weekday_fields = {
                0: "monday_stores",
                1: "tuesday_stores",
                2: "wednesday_stores",
                3: "thursday_stores",
                4: "friday_stores",
                5: "saturday_stores",
                6: "sunday_stores",
            }

            field_name = weekday_fields.get(weekday)
            if not field_name:
                return Response(
                    {
                        "success": False,
                        "error": "Неверный день недели",
                    },
                    status=status.HTTP_400_BAD_REQUEST,
                )

            created_plans = []

            for agent in Agent.objects.all():
                planned_stores = getattr(agent, field_name).all()
                planned_count = planned_stores.count()

                visited_stores_count = (
                    Store.objects.filter(posts__agent=agent, posts__created__date=today)
                    .distinct()
                    .count()
                )

                daily_plan, created = DailyPlan.objects.update_or_create(
                    agent=agent,
                    date=today,
                    defaults={
                        "planned_stores_count": planned_count,
                        "visited_stores_count": visited_stores_count,
                    },
                )

                created_plans.append(
                    {
                        "agent_id": agent.id,
                        "agent_number": agent.agent_number,
                        "planned_stores_count": planned_count,
                        "visited_stores_count": visited_stores_count,
                        "created": created,
                    }
                )

            return Response(
                {
                    "success": True,
                    "message": "Ежедневные планы успешно записаны",
                    "date": today,
                    "plans_created": len(created_plans),
                    "details": created_plans,
                },
                status=status.HTTP_201_CREATED,
            )

        except Exception as e:
            return Response(
                {
                    "success": False,
                    "error": f"Ошибка при записи планов: {str(e)}",
                },
                status=status.HTTP_500_INTERNAL_SERVER_ERROR,
            )
