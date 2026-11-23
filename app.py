            elif category:
                totals[category] += amount

        internal = totals["Внутренние услуги"]
        shop = totals["Магазинчик"]
        equipment = totals["Оборудование"]
        consumables = totals["Расходники"]
        services = totals["Услуги"]
        publications = totals["Публикации"]
        travel = totals["Командировки"]
        equipment_ocdi = totals["Оборудование (ОЦДИ)"]

        # В бюджет входят все рассчитанные суммы, кроме внутренних услуг,
        # магазинчика и оборудования с ОЦДИ.
        budget = sum(
            amount for category, amount in totals.items()
            if category not in {"Внутренние услуги", "Магазинчик", "Оборудование (ОЦДИ)"}
        )

        title_rows.append(
            project["title"] + [
                budget, None, None, internal, shop, equipment,
                consumables, services, publications, travel,
                equipment_ocdi, project["filename"]
            ]
        )
        team_rows.extend(project["team"])
        purchase_rows.extend(project["purchases"])

    # Суммируем R по одинаковым полноценным ФИО между всеми загруженными проектами.
    employee_totals = defaultdict(float)
    employee_files = defaultdict(list)
    employee_ids = defaultdict(list)
    for row in team_rows:
        name, share = row[0], row[17]
        if is_full_person_name(name) and isinstance(share, (int, float)):
            key = re.sub(r"\s+", " ", name.strip()).casefold()
            employee_totals[key] += share
            employee_files[key].append(row[19])
            employee_ids[key].append(row[18])

    for key, total in employee_totals.items():
        if total > 1 + 1e-9:
            name = next(row[0] for row in team_rows if is_full_person_name(row[0]) and
                        re.sub(r"\s+", " ", row[0].strip()).casefold() == key)
            files = ", ".join(dict.fromkeys(employee_files[key]))
            ids = ", ".join(dict.fromkeys(str(x) for x in employee_ids[key] if x))
            add_issue(logs, "Предупреждение", files, ids, TEAM_SHEET, "R",
                      name, "Суммарная доля рабочего времени по проектам превышает 1.",
                      total, total, "Уточнить распределение занятости сотрудника.")

    output = Workbook()
    output.remove(output.active)

    sheets = [
        ("Титульные", TITLE_HEADERS, title_rows),
        ("Команда", TEAM_HEADERS, team_rows),
        ("Закупки", PURCHASE_HEADERS, purchase_rows),
        ("Журнал ошибок", LOG_HEADERS, logs),
    ]

    for sheet_name, headers, rows in sheets:
        ws = output.create_sheet(sheet_name)
        ws.append(headers)
        for row in rows:
            ws.append(list(row))
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 34
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="244062")
            cell.alignment = Alignment(vertical="center", wrap_text=True)
        for col_idx in range(1, ws.max_column + 1):
            letter = get_column_letter(col_idx)
            sample = [
                len(str(ws.cell(r, col_idx).value or ""))
                for r in range(1, min(ws.max_row, 150) + 1)
            ]
            ws.column_dimensions[letter].width = min(42, max(12, max(sample, default=12) + 2))

    if not logs:
        ws = output["Журнал ошибок"]
        ws.append(["Нет ошибок", None, None, None, None, None,
                   "Проверка завершена", "Ошибок и предупреждений не обнаружено."])

    buffer = io.BytesIO()
    output.save(buffer)
    return buffer.getvalue()


# -------------------- ИНТЕРФЕЙС --------------------

if build_clicked and uploaded_files:
        issues = []
        projects = []
        progress = st.progress(0)
        status = st.empty()

        for index, uploaded in enumerate(uploaded_files, start=1):
            status.write(f"Обрабатывается {uploaded.name} ({index}/{len(uploaded_files)})")
            try:
                project = read_project(uploaded, issues)
                if project is not None:
                    projects.append(project)
            except Exception as exc:
                add_issue(
                    issues, "Ошибка", uploaded.name, None, "Обработка файла", "-",
                    "Непредвиденная ошибка", repr(exc), traceback.format_exc(),
                    None, "Файл пропущен; остальные файлы продолжают обрабатываться."
                )
            progress.progress(index / len(uploaded_files))

        try:
            excel_bytes = build_output(projects, issues)
            error_count = sum(1 for item in issues if item[0] == "Ошибка")
            warning_count = sum(1 for item in issues if item[0] == "Предупреждение")

            st.success(f"Сборка завершена: обработано проектов {len(projects)} из {len(uploaded_files)}.")
            col1, col2, col3 = st.columns(3)
            col1.metric("Проектов", len(projects))
            col2.metric("Ошибок", error_count)
            col3.metric("Предупреждений", warning_count)

            st.download_button(
                "⬇ Скачать итоговый Excel",
                data=excel_bytes,
                file_name="Итоговая_сборка_бюджета.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            if issues:
                st.subheader("Журнал проверки")
                st.dataframe(
                    pd.DataFrame(issues, columns=LOG_HEADERS),
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.success("Ошибок и предупреждений не обнаружено.")
        except Exception as exc:
            st.error("Не удалось сформировать итоговый Excel.")
            st.code(traceback.format_exc())
