
import time
import pyautogui


class ImageOperation(object):

    @staticmethod
    def loop_retry_check(loop_search_image):
        print(loop_search_image)
        loop_check_1 = pyautogui.locateOnScreen(loop_search_image, confidence=0.9, grayscale=True)
        loop_check_2 = pyautogui.locateOnScreen(loop_search_image, confidence=0.9, grayscale=True)
        while loop_check_1 is None and loop_check_2 is None:
            print("Image Loop Check ::")
            print("Check_1", loop_check_1)
            print("Check_2", loop_check_2)
            time.sleep(2)
            loop_check_1 = pyautogui.locateOnScreen(loop_search_image, confidence=0.9, grayscale=True)
            loop_check_2 = pyautogui.locateOnScreen(loop_search_image, confidence=0.9, grayscale=True)
        print("Image Loop Check ::")
        print("Final_1", loop_check_1)
        print("Final_2", loop_check_2)
        return loop_check_1, loop_check_2

    @staticmethod
    def image_check(search_image):
        print(search_image)
        check_1 = pyautogui.locateOnScreen(search_image, confidence=0.9, grayscale=True)
        check_2 = pyautogui.locateOnScreen(search_image, confidence=0.9, grayscale=True)
        print("Image Check ::")
        print("Check_1", check_1)
        print("Check_2", check_2)
        return check_1, check_2
