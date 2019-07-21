from bisect import bisect_left
from time import sleep

import numpy as np
from cv2 import VideoCapture, CAP_DSHOW
from win32com.client import GetObject


class Autobrightness:
    WMI = GetObject('winmgmts:\\\\.\\root\\WMI')

    @staticmethod
    def _get_image():
        cam = VideoCapture(0, CAP_DSHOW)
        result_code, image = cam.read()
        cam.release()
        if result_code:
            return image

    @staticmethod
    def _get_average_rgb(image):
        return np.mean(np.flip(image, axis=2), axis=(0, 1))

    def _calculate_brightness(self, image):
        r, g, b = self._get_average_rgb(image)
        return int(0.2126*r + 0.7152*g + 0.0722*b)

    @staticmethod
    def _get_closest_brightness_level(rgb_lvls, brightness_of_image):
        return bisect_left(rgb_lvls, brightness_of_image)

    def _get_brightness_api(self):
        return self.WMI.InstancesOf('WmiMonitorBrightnessMethods')[0]

    def _get_current_brightness(self):
        obj = self.WMI.InstancesOf('WmiMonitorBrightness')[0]
        return obj.CurrentBrightness

    def _set_brightness(self, *args):
        api = self._get_brightness_api()
        method = api.Methods_('WmiSetBrightness')
        parameters = method.InParameters
        for i, arg in enumerate(args):
            parameters.Properties_[i].Value = arg
        api.ExecMethod_('WmiSetBrightness', parameters)

    def run(self):
        rgb_levels = list(range(10, 251, 24))  # ten rgb levels
        brightness_levels = list(range(0, 101, 10))  # ten levels of brightness
        old_brightness = self._get_current_brightness()
        while True:
            current_brightness = self._get_current_brightness()
            if current_brightness == old_brightness:  # check if brightness was adjusted manually, break if it was
                id_ = self._get_closest_brightness_level(rgb_levels, self._calculate_brightness(self._get_image()))
                new_brightness = brightness_levels[id_]
                if current_brightness != new_brightness:
                    self._set_brightness(new_brightness, 0)
                    old_brightness = new_brightness
                else:
                    sleep(5.0)
            else:
                break


autobrightness = Autobrightness()
autobrightness.run()
