#!/usr/bin/env python3

import cv2
import numpy as np

bottom_margin = 62  # in px

scale = 20  # in um

nb_pixels_scale = 428  # in px, 428 px = 20 um


def count_droplets(image_path, debug=False):

    image_extension = image_path.split(".")[-1]
    if image_extension == "tif":
        img = cv2.imread(image_path, cv2.IMREAD_UNCHANGED)
    elif image_extension == "jpg":
        img = cv2.imread(image_path)

    image_name = image_path.split("/")[-1].split("\\")[-1]

    cropped_img = img[0 : img.shape[0] - bottom_margin, :]

    if debug:
        cv2.imshow(f"original img ({image_name})", cropped_img)

    if len(img.shape) == 3:
        img_gray = cv2.cvtColor(cropped_img, cv2.COLOR_BGR2GRAY)
    else:
        img_gray = cropped_img

    ret, thresh = cv2.threshold(img_gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

    contours, hierarchy = cv2.findContours(
        thresh.copy(), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE
    )

    img_contours = cv2.cvtColor(cropped_img, cv2.COLOR_BGR2RGB)

    droplet_diameters = []
    for cont in contours:
        (cx, cx), (w, h), a = cv2.minAreaRect(cont)
        aspect_ratio = min(w, h) / max(w, h, 1)
        area = cv2.contourArea(cont)
        if aspect_ratio > 0.4 and area > 0:
            equi_diameter = np.sqrt(4 * area / np.pi)
            diameter_in_um = equi_diameter / nb_pixels_scale * scale
            droplet_diameters.append(diameter_in_um)
            if debug:
                cv2.drawContours(img_contours, cont, -1, (0, 0, 255), 2)
            else:
                cv2.drawContours(img_contours, cont, -1, (255, 0, 0), 2)
        elif debug:
            cv2.drawContours(img_contours, cont, -1, (255, 0, 0), 2)

    if debug:
        cv2.imshow(f"contours ({image_name})", img_contours)

    if debug:
        while True:
            k = cv2.waitKey(0)
            if k == 120:  # close if 'x' key pressed
                cv2.destroyAllWindows()
                break

    return droplet_diameters, cropped_img, img_contours
