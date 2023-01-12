from django.http.response import HttpResponse

from rest_framework.viewsets import ModelViewSet
from rest_framework.permissions import IsAuthenticated
from rest_framework.decorators import action
from rest_framework.response import Response

from app.models import AppModel
from app.serializers import AppSerializer
from app.helpers import genrate_model_fields, genrate_dynamic_excel_data, generate_excel


class AppViewSet(ModelViewSet):
    permission_classes = [IsAuthenticated]
    serializer_class = AppSerializer

    def get_queryset(self):
        queryset = AppModel.objects.all().order_by("-id")
        return queryset

    @action(detail=False, methods=["GET"])
    def dynamic_excel_report(self, request):
        fields = self.request.GET.get("fields", None)

        if not fields:
            return Response("FIELDS IS REQUIRED!", status=400)

        data = self.get_queryset()

        excel = genrate_dynamic_excel_data(fields, data)

        response = HttpResponse(content_type="application/ms-excel")
        response["Content-Disposition"] = (
            "attachment;filename="
            + str(data.first().__class__.__name__)
            + "-Excel-Report.csv"
        )
        generate_excel(response, excel)
        return response

    @action(detail=False, methods=["GET"])
    def get_fields(self, request):
        model_class = self.get_queryset().first().__class__

        fields = genrate_model_fields(model_class)

        return Response(fields)
