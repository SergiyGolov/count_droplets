#!/usr/bin/env python3

import matplotlib.pyplot as plt
import numpy as np

bins = [0, 0.15, 0.6, 2, 2.1]


def plot_bins(data, title, image_paths):
    clipped_data = [np.clip(x, 0, 2.00001) for x in data]
    fig = plt.figure(f"Droplet diameter histogram for {title}")
    bar_width = 1 / len(data)
    histograms = []
    for i, (x, image_path) in enumerate(zip(clipped_data, image_paths)):
        h, e = np.histogram(x, bins=bins)
        histograms.append(h)
        if len(data) > 1:
            plt.bar(
                [y + i * bar_width for y in range(len(bins) - 1)],
                h,
                width=bar_width,
                edgecolor="k",
                label=image_path,
            )
        else:
            plt.bar(
                [y for y in range(len(bins) - 1)], h, width=bar_width, edgecolor="k"
            )
    # plt.hist(clipped_data, bins=bins, edgecolor="black",label=image_paths)

    if len(data) > 1:
        plt.legend(loc="upper right")

    plt.xticks(
        ticks=[x - bar_width / 2 for x in range(len(bins))], labels=bins[:-1] + [">2"]
    )
    plt.xlabel("Diameter (µm)")
    plt.ylabel("Number of droplets")
    plt.title(f"Droplet diameter histogram for\n{title}")
    plt.show(block=False)

    if len(data) > 1:
        for i, (b1, b2) in enumerate(zip(bins[:-1], bins[1:-1] + [">2"])):

            if b1 == bins[-2]:
                fig_title = f"Comparing configurations for droplets of size {b2} µm"
            else:
                fig_title = (
                    f"Comparing configurations for droplets of size {b1} - {b2} µm"
                )

            fig = plt.figure(fig_title)
            plt.title(fig_title)
            plt.xlabel("Configuration")
            plt.ylabel("Number of droplets")

            plt.xticks(ticks=[x for x in range(len(image_paths))], labels=image_paths)

            h = [histograms[x][i] for x in range(len(image_paths))]

            plt.bar(
                range(len(image_paths)), h, width=0.5, edgecolor="k", label=image_path
            )

            plt.show(block=False)

    plt.show()
