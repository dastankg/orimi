from datetime import datetime, timedelta

import openpyxl
from django.http import HttpResponse
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .models import DailyPlan, PhotoPost


def export_to_excel(modeladmin, request, queryset):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Отчет фотопостов"

    end_date = datetime.now()
    start_date = end_date - timedelta(days=6)

    selected_agents = queryset

    # Стили
    header_font = Font(bold=True, size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    link_font = Font(color="0000FF", underline="single")

    dmp_orimi_brands = ["Гринф", "Tess", "ЖН", "Шах", "Jardin", "Жокей"]
    dmp_competitor_brands = ["Бета", "Пиала", "Ахмад", "Jacobs", "Nestlе"]
    all_dmp_brands = dmp_orimi_brands + dmp_competitor_brands

    brand_mapping = {
        "гринф": ["гринф", "greenf", "grinf"],
        "tess": ["tess", "тесс"],
        "жн": ["жн", "jn"],
        "шах": ["шах", "shah"],
        "jardin": ["jardin", "жардин"],
        "жокей": ["жокей", "jockey"],
        "бета": ["бета", "beta"],
        "пиала": ["пиала", "piala"],
        "ахмад": ["ахмад", "ahmad"],
        "jacobs": ["jacobs", "якобс"],
        "nestlе": ["nestle", "нестле", "nestlе"],
    }

    def check_brand_match(dmp_type, brand, brand_variants):
        if not dmp_type:
            return False
        dmp_type_lower = dmp_type.strip().lower()

        if dmp_type_lower == brand.lower():
            return True

        return dmp_type_lower in brand_variants

    row = 1

    main_headers = [
        "Дата",
        "SNO",
        "Наименование TT",
        "Начала работы в TT",
        "Завершение работы в TT",
        "Time in outlet",
        "РМП_чай_ДО",
        "РМП_чай_ПОСЛЕ",
        "РМП_кофе_ДО",
        "РМП_кофе_ПОСЛЕ",
        "ДМП_ОРИМИ КР",
        "ДМП_конкурент",
    ]

    col = 1
    for i, header in enumerate(main_headers[:10]):
        ws.cell(row=row, column=col + i, value=header).font = header_font
        ws.merge_cells(
            start_row=row, start_column=col + i, end_row=row + 1, end_column=col + i
        )
        ws.cell(row=row, column=col + i).alignment = center_align

    col_offset = 11
    ws.cell(row=row, column=col_offset, value="ДМП_ОРИМИ КР").font = header_font
    ws.merge_cells(
        start_row=row,
        start_column=col_offset,
        end_row=row,
        end_column=col_offset + len(dmp_orimi_brands) - 1,
    )
    ws.cell(row=row, column=col_offset).alignment = center_align

    col_offset += len(dmp_orimi_brands)
    ws.cell(row=row, column=col_offset, value="ДМП_конкурент").font = header_font
    ws.merge_cells(
        start_row=row,
        start_column=col_offset,
        end_row=row,
        end_column=col_offset + len(dmp_competitor_brands) - 1,
    )
    ws.cell(row=row, column=col_offset).alignment = center_align

    col = 11
    for brand in all_dmp_brands:
        ws.cell(row=row + 1, column=col, value=brand).font = header_font
        ws.cell(row=row + 1, column=col).alignment = center_align
        col += 1

    data_row = 3

    for agent in selected_agents:
        agent_has_data = False

        for day_offset in range(7):
            current_date = start_date + timedelta(days=day_offset)
            date_str = current_date.strftime("%d.%m.%Y")

            day_posts = (
                PhotoPost.objects.filter(agent=agent, created__date=current_date.date())
                .select_related("store")
                .order_by("created")
            )

            if not day_posts.exists():
                continue

            agent_has_data = True

            date_start_row = data_row

            store_groups = []
            current_store = None
            current_group = []

            for post in day_posts:
                if post.store != current_store:
                    if current_group:
                        store_groups.append((current_store, current_group))
                    current_store = post.store
                    current_group = [post]
                else:
                    current_group.append(post)

            if current_group:
                store_groups.append((current_store, current_group))

            first_store_in_day = True

            for store, posts_list in store_groups:
                rmp_photos = {}
                for post_type in [
                    "РМП_чай_ДО",
                    "РМП_чай_ПОСЛЕ",
                    "РМП_кофе_ДО",
                    "РМП_кофе_ПОСЛЕ",
                ]:
                    type_posts = [
                        p for p in posts_list if p.post_type == post_type and p.image
                    ]
                    if type_posts:
                        rmp_photos[post_type] = type_posts

                max_photos = (
                    max(len(photos) for photos in rmp_photos.values()) if rmp_photos else 1
                )

                for photo_index in range(max_photos):
                    if photo_index == 0:
                        if first_store_in_day:
                            ws.cell(row=data_row, column=1, value=date_str)
                            ws.cell(row=data_row, column=2, value=agent.agent_name)
                            first_store_in_day = False

                        ws.cell(row=data_row, column=3, value=store.name)

                        first_post = min(posts_list, key=lambda x: x.created)
                        last_post = max(posts_list, key=lambda x: x.created)

                        ws.cell(
                            row=data_row,
                            column=4,
                            value=first_post.created.strftime("%H:%M"),
                        )
                        ws.cell(
                            row=data_row,
                            column=5,
                            value=last_post.created.strftime("%H:%M"),
                        )

                        time_diff = last_post.created - first_post.created
                        hours = int(time_diff.total_seconds() // 3600)
                        minutes = int((time_diff.total_seconds() % 3600) // 60)
                        ws.cell(row=data_row, column=6, value=f"{hours}:{minutes:02d}")

                    col = 7
                    for post_type in [
                        "РМП_чай_ДО",
                        "РМП_чай_ПОСЛЕ",
                        "РМП_кофе_ДО",
                        "РМП_кофе_ПОСЛЕ",
                    ]:
                        if post_type in rmp_photos and photo_index < len(
                            rmp_photos[post_type]
                        ):
                            post = rmp_photos[post_type][photo_index]
                            image_url = request.build_absolute_uri(post.image.url)
                            ws.cell(
                                row=data_row,
                                column=col,
                                value=f'=HYPERLINK("{image_url}","фото")',
                            )
                            ws.cell(row=data_row, column=col).font = link_font
                        else:
                            ws.cell(row=data_row, column=col, value="")
                        col += 1

                    if photo_index == 0:
                        col = 11

                        dmp_orimi_posts = [
                            p for p in posts_list if p.post_type == "ДМП_ОРИМИ КР"
                        ]
                        for brand in dmp_orimi_brands:
                            brand_variants = brand_mapping.get(
                                brand.lower(), [brand.lower()]
                            )

                            brand_posts = []
                            for post in dmp_orimi_posts:
                                if check_brand_match(post.dmp_type, brand, brand_variants):
                                    brand_posts.append(post)

                            if brand_posts and brand_posts[0].image:
                                image_url = request.build_absolute_uri(
                                    brand_posts[0].image.url
                                )
                                ws.cell(
                                    row=data_row,
                                    column=col,
                                    value=f'=HYPERLINK("{image_url}","фото")',
                                )
                                ws.cell(row=data_row, column=col).font = link_font
                            else:
                                ws.cell(row=data_row, column=col, value="")
                            col += 1

                        dmp_competitor_posts = [
                            p for p in posts_list if p.post_type == "ДМП_конкурент"
                        ]
                        for brand in dmp_competitor_brands:
                            brand_variants = brand_mapping.get(
                                brand.lower(), [brand.lower()]
                            )

                            brand_posts = []
                            for post in dmp_competitor_posts:
                                if check_brand_match(post.dmp_type, brand, brand_variants):
                                    brand_posts.append(post)

                            if brand_posts:
                                total_count = sum(p.dmp_count or 1 for p in brand_posts)
                                ws.cell(row=data_row, column=col, value=total_count)
                            else:
                                ws.cell(row=data_row, column=col, value="")
                            col += 1

                    data_row += 1

            if date_start_row < data_row:
                ws.merge_cells(
                    start_row=date_start_row,
                    start_column=1,
                    end_row=data_row - 1,
                    end_column=1,
                )
                ws.merge_cells(
                    start_row=date_start_row,
                    start_column=2,
                    end_row=data_row - 1,
                    end_column=2,
                )

        if agent_has_data:
            data_row += 1

    for row in ws.iter_rows(
        min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
    ):
        for cell in row:
            cell.alignment = center_align
            cell.border = thin_border

    # Настройка ширины колонок
    column_widths = {
        1: 12,  # Дата
        2: 15,  # SNO (Агент)
        3: 25,  # Магазин
        4: 15,  # Начало
        5: 15,  # Конец
        6: 12,  # Time in outlet
        7: 15,  # РМП_чай_ДО
        8: 15,  # РМП_чай_ПОСЛЕ
        9: 15,  # РМП_кофе_ДО
        10: 15,  # РМП_кофе_ПОСЛЕ
    }

    for col_num, width in column_widths.items():
        column_letter = get_column_letter(col_num)
        ws.column_dimensions[column_letter].width = width

    for col_num in range(11, 11 + len(all_dmp_brands)):
        column_letter = get_column_letter(col_num)
        ws.column_dimensions[column_letter].width = 10

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    agent_names = "_".join(
        [agent.agent_name[:10].replace(" ", "_") for agent in selected_agents[:3]]
    )
    if len(selected_agents) > 3:
        agent_names += f"_and_{len(selected_agents) - 3}_more"

    filename = f"photo_report_{agent_names}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    response["Content-Disposition"] = f"attachment; filename={filename}"

    wb.save(response)

    modeladmin.message_user(
        request, f"Отчет успешно создан для {len(selected_agents)} агент(ов)"
    )

    return response


export_to_excel.short_description = "Выгрузить отчет в Excel"


def export_plan_visits_to_excel(modeladmin, request, queryset):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "План визитов"

    selected_agents = queryset

    header_font = Font(bold=True, size=10)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )
    success_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")
    fail_fill = PatternFill(start_color="FFB6C1", end_color="FFB6C1", fill_type="solid")
    header_fill = PatternFill(start_color="D3D3D3", end_color="D3D3D3", fill_type="solid")

    dmp_orimi_brands = ["Гринф", "Tess", "ЖН", "Шах", "Jardin", "Жокей"]
    dmp_competitor_brands = ["Бета", "Пиала", "Ахмад", "Jacobs", "Nestlе"]
    all_dmp_brands = dmp_orimi_brands + dmp_competitor_brands

    brand_variations = {
        "Бета": ["бета", "beta", "Beta", "BETA"],
        "Пиала": ["пиала", "piala", "Piala", "PIALA"],
        "Ахмад": ["ахмад", "ahmad", "Ahmad", "AHMAD"],
        "Гринф": ["гринф", "greenf", "Greenf", "GREENF", "grinf", "Grinf"],
        "Tess": ["tess", "Tess", "TESS", "тесс", "Тесс"],
        "ЖН": ["жн", "ЖН", "jn", "JN"],
        "Шах": ["шах", "shah", "Shah", "SHAH"],
        "Jardin": ["jardin", "Jardin", "JARDIN", "жардин", "Жардин"],
        "Жокей": ["жокей", "jockey", "Jockey", "JOCKEY"],
        "Jacobs": ["jacobs", "Jacobs", "JACOBS", "якобс", "Якобс"],
        "Nestlе": ["nestle", "Nestle", "NESTLE", "нестле", "Нестле"],
    }

    headers_row1 = [
        ("Дата", 1, 1),
        ("Мерчендайзер", 2, 2),
        ("План визита", 3, 3),
        ("Факт визита", 4, 4),
        ("Call completion %", 5, 5),
        ("Кол-во Визитов < 30 мин", 6, 6),
        ("Кол-во Визитов > 30 мин", 7, 7),
        ("Общая Время в ТТ", 8, 8),
        ("РМП_чай", 9, 10),
        ("РМП_кофе", 11, 12),
        ("ДМП ОРИМИ", 13, 14),
    ]

    current_col = 15
    for brand in all_dmp_brands:
        headers_row1.append((brand, current_col, current_col))
        current_col += 1

    headers_row2 = [
        (9, "Кол-во ТТ"),
        (10, "кол-во фото"),
        (11, "Кол-во ТТ"),
        (12, "кол-во фото"),
        (13, "Кол-во ТТ"),
        (14, "сумма"),
    ]

    for header, start_col, end_col in headers_row1:
        ws.cell(row=1, column=start_col, value=header).font = header_font
        ws.cell(row=1, column=start_col).fill = header_fill
        if start_col != end_col:
            ws.merge_cells(
                start_row=1, start_column=start_col, end_row=1, end_column=end_col
            )
            ws.cell(row=1, column=end_col).fill = header_fill

    for col, header in headers_row2:
        ws.cell(row=2, column=col, value=header).font = header_font
        ws.cell(row=2, column=col).fill = header_fill

    single_level_headers = list(range(1, 9)) + list(range(15, 15 + len(all_dmp_brands)))
    for i in single_level_headers:
        ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)

    row = 3

    for agent in selected_agents:
        daily_plans = DailyPlan.objects.filter(agent=agent).order_by("date")

        for daily_plan in daily_plans:
            date = daily_plan.date
            plan_count = daily_plan.planned_stores_count
            fact_count = daily_plan.visited_stores_count

            visits_under_30 = 0
            visits_over_30 = 0
            total_time_minutes = 0

            visited_stores = (
                PhotoPost.objects.filter(agent=agent, created__date=date)
                .values_list("store", flat=True)
                .distinct()
            )

            for store_id in visited_stores:
                store_posts = PhotoPost.objects.filter(
                    agent=agent, store_id=store_id, created__date=date
                ).order_by("created")

                if store_posts.exists():
                    first_post = store_posts.first()
                    last_post = store_posts.last()
                    time_diff = last_post.created - first_post.created
                    minutes_in_store = int(time_diff.total_seconds() // 60)

                    if minutes_in_store < 30:
                        visits_under_30 += 1
                    else:
                        visits_over_30 += 1

                    total_time_minutes += minutes_in_store

            hours = total_time_minutes // 60
            minutes = total_time_minutes % 60
            total_time_str = f"{hours}:{minutes:02d}"

            posts_data = PhotoPost.objects.filter(agent=agent, created__date=date).values(
                "post_type", "store", "dmp_type", "dmp_count"
            )

            rmp_tea_stores = set()
            rmp_tea_photos = 0
            rmp_coffee_stores = set()
            rmp_coffee_photos = 0
            dmp_orimi_stores = set()
            dmp_orimi_sum = 0
            dmp_competitor_stores = set()
            dmp_competitor_sum = 0

            brand_sums = dict.fromkeys(all_dmp_brands, 0)

            for post in posts_data:
                post_type = post["post_type"]
                store_id = post["store"]
                dmp_type = post["dmp_type"] or ""
                dmp_count = post["dmp_count"] or 0

                if "РМП_чай" in post_type:
                    rmp_tea_stores.add(store_id)
                    rmp_tea_photos += 1
                elif "РМП_кофе" in post_type:
                    rmp_coffee_stores.add(store_id)
                    rmp_coffee_photos += 1
                elif "ДМП_ОРИМИ" in post_type:
                    dmp_orimi_stores.add(store_id)
                    dmp_orimi_sum += dmp_count
                    for brand in dmp_orimi_brands:
                        brand_variants = brand_variations.get(brand, [brand.lower()])
                        for variant in brand_variants:
                            if variant in dmp_type.lower():
                                brand_sums[brand] += dmp_count
                                break
                elif "ДМП_конкурент" in post_type:
                    dmp_competitor_stores.add(store_id)
                    dmp_competitor_sum += dmp_count
                    for brand in dmp_competitor_brands:
                        brand_variants = brand_variations.get(brand, [brand.lower()])
                        for variant in brand_variants:
                            if variant in dmp_type.lower():
                                brand_sums[brand] += dmp_count
                                break

            ws.cell(row=row, column=1, value=date.strftime("%d.%m"))
            ws.cell(row=row, column=2, value=agent.agent_name)
            ws.cell(row=row, column=3, value=plan_count)
            ws.cell(row=row, column=4, value=fact_count)

            # Call completion %
            if plan_count > 0:
                completion_percent = round((fact_count / plan_count) * 100)
                ws.cell(row=row, column=5, value=f"{completion_percent}%")
                if completion_percent >= 80:
                    ws.cell(row=row, column=5).fill = success_fill
                else:
                    ws.cell(row=row, column=5).fill = fail_fill
            else:
                ws.cell(row=row, column=5, value="0%")
                ws.cell(row=row, column=5).fill = PatternFill(
                    start_color="CCCCCC", end_color="CCCCCC", fill_type="solid"
                )

            ws.cell(row=row, column=6, value=visits_under_30)
            ws.cell(row=row, column=7, value=visits_over_30)
            ws.cell(row=row, column=8, value=total_time_str)

            ws.cell(row=row, column=9, value=len(rmp_tea_stores))
            ws.cell(row=row, column=10, value=rmp_tea_photos)
            ws.cell(row=row, column=11, value=len(rmp_coffee_stores))
            ws.cell(row=row, column=12, value=rmp_coffee_photos)
            ws.cell(row=row, column=13, value=len(dmp_orimi_stores))
            ws.cell(row=row, column=14, value=dmp_orimi_sum)  # Сумма для ДМП ОРИМИ

            col = 15
            for brand in all_dmp_brands:
                ws.cell(row=row, column=col, value=brand_sums[brand])
                col += 1

            row += 1

    for row in ws.iter_rows(
        min_row=1, max_row=ws.max_row, min_col=1, max_col=ws.max_column
    ):
        for cell in row:
            cell.alignment = center_align
            cell.border = thin_border

    # Настройка ширины колонок
    column_widths = {
        "A": 12,  # Дата
        "B": 20,  # Мерчендайзер
        "C": 12,  # План визита
        "D": 12,  # Факт визита
        "E": 15,  # Call completion %
        "F": 18,  # Визитов < 30 мин
        "G": 18,  # Визитов > 30 мин
        "H": 15,  # Общее время
        "I": 10,  # РМП чай ТТ
        "J": 12,  # РМП чай фото
        "K": 10,  # РМП кофе ТТ
        "L": 12,  # РМП кофе фото
        "M": 12,  # ДМП ОРИМИ ТТ
        "N": 12,  # ДМП ОРИМИ сумма
    }

    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    for i in range(15, 15 + len(all_dmp_brands)):
        col_letter = get_column_letter(i)
        ws.column_dimensions[col_letter].width = 10

    response = HttpResponse(
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    filename = f"plan_visits_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    response["Content-Disposition"] = f"attachment; filename={filename}"

    wb.save(response)

    modeladmin.message_user(
        request, f"Отчет по плану визитов создан для {len(selected_agents)} агент(ов)"
    )

    return response


export_plan_visits_to_excel.short_description = "Выгрузить план визитов в Excel"
