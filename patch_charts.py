import sys

def patch():
    with open('generate_monthly_report.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 1. Monthly Downtime Chart
    old_m1 = """    # 1. Monthly Downtime Chart (3D)
    dt_chart = BarChart3D()
    dt_chart.type = "col"
    dt_chart.grouping = "clustered"
    dt_chart.title = f"Время остановок (ч) - {get_month_name_ru(month)} {year}"
    dt_chart.title.layout = Layout(manualLayout=ManualLayout(xMode='edge', yMode='edge', x=0.4402, y=0.0))
    dt_chart.legend = None
    dt_chart.width = 15
    dt_chart.height = 7.5
    
    dt_data = Reference(ws, min_col=2, min_row=summary_header_row + 1, max_row=summary_end_row)
    dt_cats = Reference(ws, min_col=4, min_row=summary_header_row + 1, max_row=summary_end_row) # Use combined labels
    dt_chart.add_data(dt_data, titles_from_data=False)
    dt_chart.set_categories(dt_cats)
    
    dt_chart.x_axis.delete = False
    dt_chart.y_axis.delete = False
    
    if dt_chart.series:
        from openpyxl.drawing.colors import ColorChoice
        s = dt_chart.series[0]
        s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr="FFC000"))"""

    new_m1 = """    # 1. Monthly Downtime Chart (3D)
    dt_chart = BarChart3D()
    dt_chart.type = "col"
    dt_chart.grouping = "clustered"
    dt_chart.title = f"Время остановок (ч) - {get_month_name_ru(month)} {year}"
    dt_chart.title.layout = Layout(manualLayout=ManualLayout(xMode='edge', yMode='edge', x=0.4402, y=0.0))
    dt_chart.legend = None
    dt_chart.width = 15
    dt_chart.height = 7.5
    
    dt_data = Reference(ws, min_col=2, min_row=summary_header_row + 1, max_row=summary_end_row)
    dt_cats = Reference(ws, min_col=4, min_row=summary_header_row + 1, max_row=summary_end_row) # Use combined labels
    dt_chart.add_data(dt_data, titles_from_data=False)
    dt_chart.set_categories(dt_cats)
    
    dt_chart.x_axis.delete = False
    dt_chart.y_axis.delete = False
    dt_chart.dLbls = DataLabelList()
    dt_chart.dLbls.showCatName = True
    
    if dt_chart.series:
        from openpyxl.drawing.colors import ColorChoice
        s = dt_chart.series[0]
        s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr="FFC000"))"""

    content = content.replace(old_m1, new_m1)

    # 2. Monthly Stop Count Chart
    old_m2 = """    # 2. Monthly Stop Count Chart (3D)
    sc_chart = BarChart3D()
    sc_chart.type = "col"
    sc_chart.grouping = "clustered"
    sc_chart.title = f"Количество остановок - {get_month_name_ru(month)} {year}"
    sc_chart.title.layout = Layout(manualLayout=ManualLayout(xMode='edge', yMode='edge', x=0.3673, y=0.0))
    sc_chart.legend = None
    sc_chart.width = 15
    sc_chart.height = 7.5
    
    sc_data = Reference(ws, min_col=3, min_row=summary_header_row + 1, max_row=summary_end_row)
    sc_cats = Reference(ws, min_col=5, min_row=summary_header_row + 1, max_row=summary_end_row) # Use combined labels
    sc_chart.add_data(sc_data, titles_from_data=False)
    sc_chart.set_categories(sc_cats)
    
    sc_chart.x_axis.delete = False
    sc_chart.y_axis.delete = False
    
    if sc_chart.series:
        from openpyxl.drawing.colors import ColorChoice
        s = sc_chart.series[0]
        s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr="4F81BD"))"""

    new_m2 = """    # 2. Monthly Stop Count Chart (3D)
    sc_chart = BarChart3D()
    sc_chart.type = "col"
    sc_chart.grouping = "clustered"
    sc_chart.title = f"Количество остановок - {get_month_name_ru(month)} {year}"
    sc_chart.title.layout = Layout(manualLayout=ManualLayout(xMode='edge', yMode='edge', x=0.3673, y=0.0))
    sc_chart.legend = None
    sc_chart.width = 15
    sc_chart.height = 7.5
    
    sc_data = Reference(ws, min_col=3, min_row=summary_header_row + 1, max_row=summary_end_row)
    sc_cats = Reference(ws, min_col=5, min_row=summary_header_row + 1, max_row=summary_end_row) # Use combined labels
    sc_chart.add_data(sc_data, titles_from_data=False)
    sc_chart.set_categories(sc_cats)
    
    sc_chart.x_axis.delete = False
    sc_chart.y_axis.delete = False
    sc_chart.dLbls = DataLabelList()
    sc_chart.dLbls.showCatName = True
    
    if sc_chart.series:
        from openpyxl.drawing.colors import ColorChoice
        s = sc_chart.series[0]
        s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr="4F81BD"))"""

    content = content.replace(old_m2, new_m2)

    # 3. Weekly Downtime Chart
    old_w1 = """        # Weekly Downtime Chart
        w_dt_chart = BarChart3D()
        w_dt_chart.type = "col"
        w_dt_chart.grouping = "clustered"
        w_dt_chart.title = f"Время остановок (ч) - {week_num} неделю"
        w_dt_chart.legend = None
        w_dt_chart.width = 15
        w_dt_chart.height = 7.5
        
        w_dt_data = Reference(ws, min_col=2, min_row=w_header_row + 1, max_row=w_end_row)
        w_dt_cats = Reference(ws, min_col=4, min_row=w_header_row + 1, max_row=w_end_row)
        w_dt_chart.add_data(w_dt_data, titles_from_data=False)
        w_dt_chart.set_categories(w_dt_cats)
        
        if w_dt_chart.series:
            from openpyxl.drawing.colors import ColorChoice
            s = w_dt_chart.series[0]
            s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr="FFC000"))"""

    new_w1 = """        # Weekly Downtime Chart
        w_dt_chart = BarChart3D()
        w_dt_chart.type = "col"
        w_dt_chart.grouping = "clustered"
        w_dt_chart.title = f"Время остановок (ч) - {week_num} неделю"
        w_dt_chart.legend = None
        w_dt_chart.width = 15
        w_dt_chart.height = 7.5
        
        w_dt_data = Reference(ws, min_col=2, min_row=w_header_row + 1, max_row=w_end_row)
        w_dt_cats = Reference(ws, min_col=4, min_row=w_header_row + 1, max_row=w_end_row)
        w_dt_chart.add_data(w_dt_data, titles_from_data=False)
        w_dt_chart.set_categories(w_dt_cats)
        
        w_dt_chart.x_axis.delete = False
        w_dt_chart.y_axis.delete = False
        w_dt_chart.dLbls = DataLabelList()
        w_dt_chart.dLbls.showCatName = True
        
        if w_dt_chart.series:
            from openpyxl.drawing.colors import ColorChoice
            s = w_dt_chart.series[0]
            s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr="FFC000"))"""

    content = content.replace(old_w1, new_w1)

    # 4. Weekly Stop Count Chart
    old_w2 = """        # Weekly Stop Count Chart
        w_sc_chart = BarChart3D()
        w_sc_chart.type = "col"
        w_sc_chart.grouping = "clustered"
        w_sc_chart.title = f"Количество остановок - {week_num} неделю"
        w_sc_chart.legend = None
        w_sc_chart.width = 15
        w_sc_chart.height = 7.5
        
        w_sc_data = Reference(ws, min_col=3, min_row=w_header_row + 1, max_row=w_end_row)
        w_sc_cats = Reference(ws, min_col=5, min_row=w_header_row + 1, max_row=w_end_row)
        w_sc_chart.add_data(w_sc_data, titles_from_data=False)
        w_sc_chart.set_categories(w_sc_cats)
        
        if w_sc_chart.series:
            from openpyxl.drawing.colors import ColorChoice
            s = w_sc_chart.series[0]
            s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr="4F81BD"))"""

    new_w2 = """        # Weekly Stop Count Chart
        w_sc_chart = BarChart3D()
        w_sc_chart.type = "col"
        w_sc_chart.grouping = "clustered"
        w_sc_chart.title = f"Количество остановок - {week_num} неделю"
        w_sc_chart.legend = None
        w_sc_chart.width = 15
        w_sc_chart.height = 7.5
        
        w_sc_data = Reference(ws, min_col=3, min_row=w_header_row + 1, max_row=w_end_row)
        w_sc_cats = Reference(ws, min_col=5, min_row=w_header_row + 1, max_row=w_end_row)
        w_sc_chart.add_data(w_sc_data, titles_from_data=False)
        w_sc_chart.set_categories(w_sc_cats)
        
        w_sc_chart.x_axis.delete = False
        w_sc_chart.y_axis.delete = False
        w_sc_chart.dLbls = DataLabelList()
        w_sc_chart.dLbls.showCatName = True
        
        if w_sc_chart.series:
            from openpyxl.drawing.colors import ColorChoice
            s = w_sc_chart.series[0]
            s.graphicalProperties = GraphicalProperties(solidFill=ColorChoice(srgbClr="4F81BD"))"""

    content = content.replace(old_w2, new_w2)

    with open('generate_monthly_report.py', 'w', encoding='utf-8') as f:
        f.write(content)

if __name__ == '__main__':
    patch()
