import cv2
import win32gui
from window import Window
import pytesseract
from pynput.keyboard import Key, Controller
import numpy as np
import random
import time

def get_window_list():
    window_list = []
    toplist = []
    def enum_windows(hwnd, result):
        win_text = win32gui.GetWindowText(hwnd)
        window_list.append((hwnd, win_text))
    win32gui.EnumWindows(enum_windows, toplist)
    return window_list

def get_hwnd(window_list):
    game_hwnd = 0
    for (hwnd, win_text) in window_list:
        if "Wordle en Español" in win_text:
            game_hwnd = hwnd
    return game_hwnd

def get_board():
    window_list = get_window_list()
    game_hwnd = get_hwnd(window_list)
    window = Window(game_hwnd)
    board = [[("","") for _ in range(5)] for _ in range(6)]
    # TODO this path only works on Windows
    pytesseract.pytesseract.tesseract_cmd = 'A:\\Tesseract\\tesseract.exe'
   
    screenshot = window.get_screenshot()
    board_image = screenshot[50:475, 205:560]

    for i in range(len(board)):
        for j in range(len(board[0])):
            y_min, y_max = i*70, (i+1)*70-10
            x_min, x_max = j*70, (j+1)*70-10
            
            color = "gray"

            letter = board_image[y_min:y_max, x_min:x_max]
            gray = cv2.cvtColor(letter, cv2.COLOR_BGR2GRAY)
            (thresh, blackAndwhite) = cv2.threshold(gray, 254, 255, cv2.THRESH_BINARY)
            #cv2.imshow("whiteblack:", blackAndwhite)
            text = pytesseract.image_to_string(blackAndwhite, config='--psm 10 --oem 1').upper().strip()

            """screenshot = cv2.rectangle(
                screenshot,
                (205+x_min, 50+y_min),
                (205+x_max, 50+y_max),
                (0,0,255),
                2
            )"""

            letterhsv = cv2.cvtColor(letter, cv2.COLOR_BGR2HSV)
            yellow = np.array([98,176,240])
            green = np.array([71,130,184])

            if sum(sum(cv2.inRange(letterhsv, green, green))) > 7500:
                color = "green"
            elif sum(sum(cv2.inRange(letterhsv, yellow, yellow))) > 7500:
                color = "yellow"

            if text != "_":
                if len(text) > 1:
                    text = text[0]

                screenshot = cv2.putText(screenshot, text, (210+x_min, 45+y_max),
                    cv2.FONT_HERSHEY_SIMPLEX, 0.4, (255,255,255), 1)
                board[i][j] = (text, color)

    #cv2.imshow("Wordle:", screenshot)
    #cv2.imshow("Board:", board_image)

    return board

def get_word_list():
    with open("Wordle Words.txt", "r", encoding="utf8") as f:
        return f.readlines()

def get_letters(board, letters_in, letters_out, letters_pos, letters_pos_out):
    for i in range(len(board)):
        for j in range(len(board[0])):
            letter = board[i][j][0]
            color = board[i][j][1]
            
            if letter != "":
                if color == "gray" and letter not in letters_in:
                    letters_out.add(letter)
            
                if color == "green":
                    letters_in.add(letter)
                    letters_pos.add((letter, j))
                
                if color == "yellow":
                    letters_in.add(letter)
                    letters_pos_out.add((letter, j))

def get_possible_words(wordlist, letters_in, letters_out, letters_pos, letters_pos_out):
    wordlist = [word for word in wordlist if all(letter in word for letter in letters_in)]
    wordlist = [word for word in wordlist if not any(letter in word for letter in letters_out)]
    wordlist = [word for word in wordlist if all(letter == word[n] for letter, n in letters_pos)]
    wordlist = [word for word in wordlist if not any(letter == word[n] for letter, n in letters_pos_out)]
    return wordlist

def solve(word):
    keyboard.type(word)
    keyboard.press(Key.enter)
    keyboard.release(Key.enter)

if __name__ == "__main__":
    keyboard = Controller()
    wordlist = get_word_list()
    word_for_solve = []

    while True:
        board = get_board()
        
        keyboard.press(Key.enter)
        keyboard.release(Key.enter)
        
        letters_in = set([])
        letters_out = set([])
        letters_pos = set([])
        letters_pos_out = set([])

        if board[0][0][0] == "":
            keyboard.type("polen")
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)
            keyboard.type("judia")
            keyboard.press(Key.enter)
            keyboard.release(Key.enter)
            time.sleep(1)
            continue

        get_letters(
            board,
            letters_in,
            letters_out,
            letters_pos,
            letters_pos_out
        )

        possible_words = get_possible_words(
            wordlist,
            letters_in,
            letters_out,
            letters_pos,
            letters_pos_out
        )

        try:
            word_for_solve = random.choice(possible_words)
        except IndexError:
            continue
        
        solve(word_for_solve)
        time.sleep(1)

        if cv2.waitKey(1) == ord('q'):
            cv2.destroyAllWindows()
            break