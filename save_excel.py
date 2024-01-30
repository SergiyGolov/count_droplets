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
    bins=[round(3.5/10*(x+1),2) for x in range(10)]
    for i,b in enumerate(bins):
        worksheet_plot.write(i+1,0,b)

    worksheet_plot.write(0,rightmost_column_i+1,'Chart categories')

    bins_column='A'

    worksheet_plot.write_formula(1,rightmost_column_i+1,f'="0 - "&\'{worksheet_plot_name}\'!{bins_column}2&""',None,f'0 - {bins[0]}')

    for i,b in enumerate(bins[1:]):
        worksheet_plot.write_formula(i+2,rightmost_column_i+1,f'=""&\'{worksheet_plot_name}\'!{bins_column}{i+2}&" - "&\'{worksheet_plot_name}\'!{bins_column}{i+3}&""',None,f'{bins[i]} - {b}')

    worksheet_plot.write_formula(len(bins)+1,rightmost_column_i+1,f'=">"&\'{worksheet_plot_name}\'!{bins_column}{len(bins)+1}&""',None,f'>{max(bins)}')

    bins_column=get_column_from_i(rightmost_column_i)
    for i,image_path in enumerate(image_paths):
        worksheet_plot.write(0, rightmost_column_i+i+2, f'Number of droplets ({image_path})')
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
    worksheet_plot.insert_chart(f"{plot_column}1", chart, {'x_scale': 2.5, 'y_scale': 2.5})

    workbook.close()