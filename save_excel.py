#!/usr/bin/env python3
import xlsxwriter

def get_column_from_i(i):
    return chr(ord('@')+i+1)

def save_to_excel(data,image_paths,filename):
    workbook = xlsxwriter.Workbook(filename)
    worksheet_plot_name='Plot'
    worksheet_plot=workbook.add_worksheet(worksheet_plot_name)

    raw_data_worksheet_name='Droplet diameters'
    worksheet = workbook.add_worksheet(raw_data_worksheet_name)
    
    for i,image_path in enumerate(image_paths):
        worksheet.write(0, i, f'Droplet diameter [µm] ({image_path})')
    
    for i,droplet_diameters in enumerate(data):
        for j,x in enumerate(sorted(droplet_diameters,reverse=True)):
            worksheet.write(j+1, i, x)

    rightmost_column_i=0
    worksheet_plot.write(0,0,'Histogram bin limits')
    # bins=[round(3.5/10*(x+1),2) for x in range(10)]
    bins=[0.15,0.6,2]
    for i,b in enumerate(bins):
        worksheet_plot.write(i+1,0,b)

    worksheet_plot.write(0,rightmost_column_i+1,'Chart categories')

    bins_column='A'

    worksheet_plot.write_formula(1,rightmost_column_i+1,f'="0 - "&\'{worksheet_plot_name}\'!{bins_column}2&" µm"',None,f'0 - {bins[0]} µm')
    chart_categories_column=get_column_from_i(rightmost_column_i+1)

    for i,b in enumerate(bins[1:]):
        worksheet_plot.write_formula(i+2,rightmost_column_i+1,f'=""&\'{worksheet_plot_name}\'!{bins_column}{i+2}&" - "&\'{worksheet_plot_name}\'!{bins_column}{i+3}&" µm"',None,f'{bins[i]} - {b} µm')

    worksheet_plot.write_formula(len(bins)+1,rightmost_column_i+1,f'=">"&\'{worksheet_plot_name}\'!{bins_column}{len(bins)+1}&" µm"',None,f'>{max(bins)} µm')

    bins_column=get_column_from_i(rightmost_column_i)
    for i,image_path in enumerate(image_paths):
        worksheet_plot.write(0, rightmost_column_i+i+2, f'{image_path}')
        column=get_column_from_i(i)
        worksheet_plot.write_array_formula(1,rightmost_column_i+i+2,1+len(bins),rightmost_column_i+i+2,f'{{=frequency(\'{raw_data_worksheet_name}\'!{column}2:{column}{len(data[i])+1},\'{worksheet_plot_name}\'!A2:A{len(bins)+1})}}')


    chart = workbook.add_chart({"type": "column"})

    chart_title=",".join(image_paths)
    chart.set_title({"name": f'Droplet diameter histogram for {chart_title}'})
    chart.set_x_axis({"name": "Diameter (µm)"})
    chart.set_y_axis({"name": "Number of droplets"})
    
    for i,_ in enumerate(data):
        chart.add_series(
        {
            "name": [worksheet_plot_name, 0, rightmost_column_i+i+2],
            "categories": [worksheet_plot_name, 1, rightmost_column_i+1, 1+len(bins), rightmost_column_i+1],
            "values": [worksheet_plot_name, 1, rightmost_column_i+2+i, len(bins)+1 , rightmost_column_i+2+i]
        }
        )

    plot_column=get_column_from_i(3+len(data))
    worksheet_plot.insert_chart(f"{plot_column}1", chart, {'x_scale':1, 'y_scale': 1})

    if len(data)>1:
        for i in range(len(bins)+1):
            if i%2==0:
                plot_column=get_column_from_i(3+len(data))
            else:
                plot_column=get_column_from_i(3+len(data)+8)
            plot_row=int(((i-i%2)/2+1)*15+1)
            chart = workbook.add_chart({"type": "column"})
            chart.set_y_axis({"name": "Number of droplets"})
            chart.set_x_axis({"name": "Configuration"})
            if i==0:
                default_value=f"Comparing configurations for droplets of size 0 - {bins[0]} µm"
            elif i==len(bins):
                default_value=f"Comparing configurations for droplets of size >{max(bins)} µm"
            else:
                default_value=f"Comparing configurations for droplets of size {bins[i-1]} - {bins[i]} µm"
            worksheet_plot.write_formula(f"{plot_column}{plot_row}",f'="Comparing configurations for droplets of size "&{chart_categories_column}{i+2}&""',value=default_value)
            chart.set_title({"name": f'=\'{worksheet_plot_name}\'!{plot_column}{plot_row}'})

            
            chart.add_series(
            {
                "name": [worksheet_plot_name, i+1, rightmost_column_i+1],
                "categories": [worksheet_plot_name, 0, rightmost_column_i+2, 0, rightmost_column_i+2+len(data)-1],
                "values": [worksheet_plot_name, i+1, rightmost_column_i+2, i+1 , rightmost_column_i+2+len(data)-1]
            }
            )

            worksheet_plot.insert_chart(f"{plot_column}{plot_row}", chart, {'x_scale':1, 'y_scale': 1})

    workbook.close()