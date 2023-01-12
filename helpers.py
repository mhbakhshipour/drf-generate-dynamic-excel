def genrate_model_fields(model_class):
    from django.db.models.fields.related import ForeignKey, ManyToManyField

    all_fields = model_class._meta.fields + model_class._meta.many_to_many
    fields = []
    for i in all_fields:
        if i.__class__ in (ForeignKey, ManyToManyField):
            if i.related_model._meta.model_name == "user":
                for f in i.related_model._meta.fields:
                    if f.name in ["id", "full_name", "username"]:
                        f_user = (
                            ("f__" if i.__class__ is ForeignKey else "m__")
                            + i.name
                            + "__"
                            + f.name
                        )
                        fields.append(f_user)
            else:
                i_fields = (
                    i.related_model._meta.fields + i.related_model._meta.many_to_many
                )
                for f in i_fields:
                    if f.__class__ in (ForeignKey, ManyToManyField):
                        f_fields = (
                            f.related_model._meta.fields
                            + f.related_model._meta.many_to_many
                        )
                        for f_f in f_fields:
                            if f.related_model._meta.model_name == "user":
                                if f_f.name in ["id", "full_name", "username"]:
                                    f_u = (
                                        ("f__" if i.__class__ is ForeignKey else "m__")
                                        + i.name
                                        + (
                                            "__f__"
                                            if f.__class__ is ForeignKey
                                            else "__m__"
                                        )
                                        + f.name
                                        + "__"
                                        + f_f.name
                                    )
                                    fields.append(f_u)
                            else:
                                f_m = (
                                    ("f__" if i.__class__ is ForeignKey else "m__")
                                    + i.name
                                    + (
                                        "__f__"
                                        if f.__class__ is ForeignKey
                                        else "__m__"
                                    )
                                    + f.name
                                    + "__"
                                    + f_f.name
                                )
                                fields.append(f_m)
                    else:
                        f_m = (
                            ("f__" if i.__class__ is ForeignKey else "m__")
                            + i.name
                            + "__"
                            + f.name
                        )
                        fields.append(f_m)
        else:
            fields.append(i.name)

    return fields


def genrate_dynamic_excel_data(fields, data):
    header = [
        str(i.split("f__")[1]) if i.startswith("f__") else str(i)
        for i in fields.split(",")
    ]
    res = []
    res.append(header)
    for a in data:
        data = []
        for i in fields.split(","):
            if i.startswith("f__"):
                rem_start = i.split("f__")[1]
                if "__f__" in rem_start:
                    f_i_field = rem_start.split("__")[0]
                    f = rem_start.split("__f__")[1]
                    f_model = f.split("__")[0]
                    f_model_field = f.split("__")[1]
                    try:
                        z = getattr(a, f_i_field)
                        z = z.all()

                        res_str = ""
                        for o in z:
                            y = getattr(o, f_model)
                            x = getattr(y, f_model_field)
                            res_str = res_str + str(x) + ","
                        data.append(str(res_str))
                    except:
                        data.append("-")
                elif "__m__" in rem_start:
                    pass
                else:
                    try:
                        i_field = rem_start.split("__")[0]
                        i_final_field = rem_start.split("__")[1]
                        z = getattr(a, i_field)
                        x = getattr(z, i_final_field)
                        data.append(str(x))
                    except:
                        data.append("-")
            elif i.startswith("m__"):
                rem_start = i.split("m__")[1]
                if "__f__" in rem_start:
                    f_i_field = rem_start.split("__")[0]
                    f = rem_start.split("__f__")[1]
                    f_model = f.split("__")[0]
                    f_model_field = f.split("__")[1]
                    try:
                        z = getattr(a, f_i_field)
                        z = z.all()

                        res_str = ""
                        for o in z:
                            y = getattr(o, f_model)
                            x = getattr(y, f_model_field)
                            res_str = res_str + str(x) + ","
                        data.append(str(res_str))
                    except:
                        data.append("-")
                elif "__m__" in rem_start:
                    pass
                else:
                    i_m_field = rem_start.split("__")[0]
                    i_m_final_field = rem_start.split("__")[1]
                    try:
                        z = getattr(a, i_m_field)
                        z = z.all()

                        res_str = ""
                        for o in z:
                            x = getattr(o, i_m_final_field)
                            res_str = res_str + str(x) + ","
                        data.append(str(res_str))
                    except:
                        data.append("-")
            else:
                try:
                    data.append(str(getattr(a, i)))
                except:
                    data.append("-")
        res.append(data)

    return res


def generate_excel(response, context_list, data_type='normal', sheet_name="sheet"):
    import xlwt

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet(sheet_name)

    if data_type == 'list' and context_list:
        for counter, value in enumerate(context_list[0].keys()):
            ws.write(0, counter, value)

        for row_counter, row in enumerate(context_list, 1):
            for cell_counter, cell in enumerate(row.values()):
                ws.write(row_counter, cell_counter, cell)

    elif data_type == 'dict' and context_list:
        for counter, value in enumerate(context_list.keys()):
            ws.write(0, counter, value)

        for counter, value in enumerate(context_list.values(), 1):
            ws.write(1, counter, value)

    else:
        for row_counter, row_value in enumerate(context_list):
            for cell_counter, cell_value in enumerate(row_value):
                ws.write(row_counter, cell_counter, cell_value)

    return wb.save(response)
