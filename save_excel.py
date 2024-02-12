#!/usr/bin/env python3
import xlsxwriter
from xlsxwriter.utility import xl_col_to_name


def save_to_excel(data, image_paths, filename):
    workbook = xlsxwriter.Workbook(filename)

    bold_format = workbook.add_format({"bold": True})

    worksheet_plot_name = "Plot"
    worksheet_plot = workbook.add_worksheet(worksheet_plot_name)

    raw_data_worksheet_name = "Droplet diameters"
    worksheet = workbook.add_worksheet(raw_data_worksheet_name)

    for i, image_path in enumerate(image_paths):
        worksheet.write(0, i, f"Droplet diameter [µm] ({image_path})", bold_format)

    for i, droplet_diameters in enumerate(data):
        for j, x in enumerate(sorted(droplet_diameters, reverse=True)):
            worksheet.write(j + 1, i, x)

    worksheet_plot.write(0, 0, "Histogram bin limits", bold_format)

    bins = [0.15, 0.6, 2]
    for i, b in enumerate(bins):
        worksheet_plot.write(i + 1, 0, b)

    worksheet_plot.write(0, 1, "Chart categories", bold_format)

    bins_column = "A"

    worksheet_plot.write_formula(
        1,
        1,
        f'="0 - "&\'{worksheet_plot_name}\'!{bins_column}2&" µm"',
        None,
        f"0 - {bins[0]} µm",
    )
    chart_categories_column = xl_col_to_name(1)

    for i, b in enumerate(bins[1:]):
        worksheet_plot.write_formula(
            i + 2,
            1,
            f'=""&\'{worksheet_plot_name}\'!{bins_column}{i+2}&" - "&\'{worksheet_plot_name}\'!{bins_column}{i+3}&" µm"',
            None,
            f"{bins[i]} - {b} µm",
        )

    worksheet_plot.write_formula(
        len(bins) + 1,
        1,
        f'=">"&\'{worksheet_plot_name}\'!{bins_column}{len(bins)+1}&" µm"',
        None,
        f">{max(bins)} µm",
    )

    for i, image_path in enumerate(image_paths):
        worksheet_plot.write(0, i + 2, f"{image_path}", bold_format)
        column = xl_col_to_name(i)
        worksheet_plot.write_array_formula(
            1,
            i + 2,
            1 + len(bins),
            i + 2,
            f"{{=frequency('{raw_data_worksheet_name}'!{column}2:{column}{len(data[i])+1},'{worksheet_plot_name}'!A2:A{len(bins)+1})}}",
        )

    worksheet_plot.write("A7", "Average droplet diameter (µm)", bold_format)
    worksheet_plot.write("A8", "Maximum droplet diameter (µm)", bold_format)

    for i, droplet_diameters in enumerate(data):
        col = xl_col_to_name(i)
        max_row = len(droplet_diameters) + 1
        avg_droplet_diameter = sum(droplet_diameters) / len(droplet_diameters)
        max_droplet_diamter = max(droplet_diameters)
        worksheet_plot.write_formula(
            6,
            i + 2,
            f"=AVERAGE('{raw_data_worksheet_name}'!{col}2:{col}{max_row})",
            value=avg_droplet_diameter,
        )
        worksheet_plot.write_formula(
            7,
            i + 2,
            f"=MAX('{raw_data_worksheet_name}'!{col}2:{col}{max_row})",
            value=max_droplet_diamter,
        )

    chart_width_in_cells=10
    chart_height_in_cells=20

    chart = workbook.add_chart({"type": "column"})

    chart_title = ",".join(image_paths)
    chart.set_title({"name": f"Droplet diameter histogram for {chart_title}"})
    chart.set_x_axis({"name": "Diameter (µm)"})
    chart.set_y_axis({"name": "Number of droplets"})

    for i, _ in enumerate(data):
        chart.add_series(
            {
                "name": [worksheet_plot_name, 0, i + 2],
                "categories": [
                    worksheet_plot_name,
                    1,
                    1,
                    1 + len(bins),
                    1,
                ],
                "values": [
                    worksheet_plot_name,
                    1,
                    2 + i,
                    len(bins) + 1,
                    2 + i,
                ],
            }
        )

    plot_column = xl_col_to_name(3 + len(data))
    worksheet_plot.insert_chart(f"{plot_column}1", chart, {"x_scale": 1.3, "y_scale": 1.35})

    if len(data) > 1:
        for i in range(len(bins) + 1):
            if i % 2 == 0:
                plot_column = xl_col_to_name(3 + len(data))
            else:
                plot_column = xl_col_to_name(3 + len(data) + chart_width_in_cells)
            plot_row = int(((i - i % 2) / 2 + 1) * chart_height_in_cells + 1)
            chart = workbook.add_chart({"type": "column"})
            chart.set_y_axis({"name": "Number of droplets"})
            chart.set_x_axis({"name": "Configuration"})
            if i == 0:
                chart_title = (
                    f"Comparing configurations for droplets of size 0 - {bins[0]} µm"
                )
            elif i == len(bins):
                chart_title = (
                    f"Comparing configurations for droplets of size >{max(bins)} µm"
                )
            else:
                chart_title = f"Comparing configurations for droplets of size {bins[i-1]} - {bins[i]} µm"

            chart.set_title({"name": chart_title})

            chart.add_series(
                {
                    "name": [worksheet_plot_name, i + 1, 1],
                    "categories": [
                        worksheet_plot_name,
                        0,
                        2,
                        0,
                        2 + len(data) - 1,
                    ],
                    "values": [
                        worksheet_plot_name,
                        i + 1,
                        2,
                        i + 1,
                        2 + len(data) - 1,
                    ],
                }
            )
            worksheet_plot.insert_chart(
                f"{plot_column}{plot_row}", chart, {"x_scale": 1.3, "y_scale": 1.35}
            )

        chart = workbook.add_chart({"type": "column"})
        chart.set_title(
            {"name": "Comparing configurations by average droplet diameter"}
        )
        chart.set_y_axis({"name": "Average droplet diameter (µm)"})
        chart.set_x_axis({"name": "Configuration"})

        chart.add_series(
            {
                "name": [worksheet_plot_name, 6, 0],
                "categories": [
                    worksheet_plot_name,
                    0,
                    2,
                    0,
                    2 + len(data) - 1,
                ],
                "values": [
                    worksheet_plot_name,
                    6,
                    2,
                    6,
                    2 + len(data) - 1,
                ],
            }
        )

        plot_column = xl_col_to_name(3 + chart_width_in_cells + len(data))
        worksheet_plot.insert_chart(
            f"{plot_column}1", chart, {"x_scale": 1.3, "y_scale": 1.35}
        )

        chart = workbook.add_chart({"type": "column"})
        chart.set_title(
            {"name": "Comparing configurations by maximum droplet diameter"}
        )
        chart.set_y_axis({"name": "Maximum droplet diameter (µm)"})
        chart.set_x_axis({"name": "Configuration"})

        chart.add_series(
            {
                "name": [worksheet_plot_name, 7, 0],
                "categories": [
                    worksheet_plot_name,
                    0,
                    2,
                    0,
                    2 + len(data) - 1,
                ],
                "values": [
                    worksheet_plot_name,
                    7,
                    2,
                    7,
                    2 + len(data) - 1,
                ],
            }
        )

        plot_column = xl_col_to_name(3 + chart_width_in_cells*2 + len(data))
        worksheet_plot.insert_chart(
            f"{plot_column}1", chart, {"x_scale": 1.3, "y_scale": 1.35}
        )

    workbook.close()
