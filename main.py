# import necessary libraries
import pyautogui as pg
import time
from pynput.keyboard import Key, Controller
import pandas as pd

# Function to get cursor position
def get_cursor_position():
    """
    Function to prompt the user to click on the screen within 5 seconds to determine cursor position.
    """
    print("Please click on the screen within 5 seconds...")
    time.sleep(5)
    click_point = pg.position()
    mouse_x, mouse_y = click_point
    print(f"Mouse clicked at: ({mouse_x}, {mouse_y})")

# Function to simulate click on screen
def click_on_screen(mouse_x, mouse_y):
    """
    Function to simulate a click on the screen at the specified position.
    """
    pg.click(mouse_x, mouse_y)

# Function to simulate pressing tab key
def tab_key(num_of_tab):
    """
    Function to simulate pressing the tab key a specified number of times.
    """
    for tab in range(1, num_of_tab + 1):
        kb.press(Key.tab)
        kb.release(Key.tab)

# Function to move to the next line
def next_line():
    """
    Function to move to the next line by pressing the tab key twice.
    """
    tab_key(2)

# Function to add a single line in BOM
def add_bom_single_line(component_item_code, number_of_tab, quantity):
    """
    Function to add a single line in the Bill of Materials.
    """
    kb.type(component_item_code)
    time.sleep(1)
    tab_key(int(number_of_tab))
    time.sleep(1)
    kb.type(str(quantity))
    time.sleep(1)

# Function to add BOM line level
def add_bom_line_level(item_code, df):
    """
    Function to add BOM line level items.
    """
    loop_df = df[df['Bill Number'] == item_code]

    for index, value in loop_df.iterrows():
        add_bom_single_line(value['Component Item Code'], value['NumOfTab'], value['Quantity/Bill'])
        next_line()

# Function to add BOM
def add_bom(item_code, df):
    """
    Function to add Bill of Materials.
    """
    item_info = df[df['Bill Number'] == item_code]
    item_info = item_info[['Bill Number', 'CopyFrom', 'Description']]
    item_info = item_info.drop_duplicates(keep='last')
    item_info = item_info.values.tolist()
    item_code = item_info[0][0]
    ref_code = item_info[0][1]
    desc = item_info[0][2]

    click_on_screen(bill_num_pos[0], bill_num_pos[1])
    time.sleep(1)
    kb.type(item_code)
    kb.press(Key.enter)
    time.sleep(2)

    click_on_screen(crate_bill_ok_btn_pos[0], crate_bill_ok_btn_pos[1])

    time.sleep(2)
    kb.type(ref_code)
    kb.press(Key.enter)

    next_line()

    time.sleep(1)
    kb.type(desc)

    click_on_screen(additional_tab_pos[0], additional_tab_pos[1])
    time.sleep(1)

    click_on_screen(ok_btn_pos[0], ok_btn_pos[1])
    kb.press(Key.enter)
    time.sleep(1)

    click_on_screen(line_tab_pos[0], line_tab_pos[1])
    time.sleep(1)

    add_bom_line_level(item_code, df)

    time.sleep(1)
    click_on_screen(finish_ok_btn_pos[0], finish_ok_btn_pos[1])

# Function to loop through item codes
def item_code_loop(list, df):
    """
    Function to loop through item codes and add them to the Bill of Materials.
    """
    for item in list:
        print('Adding: ', item[0])
        add_bom(item[0], df)

# Define keyboard controller
kb = Controller()

# Read BOM data from Excel file
new_bom = pd.read_excel('bom_component_list.xlsx', sheet_name='Sheet1')

# Group BOM data by Bill Number
new_bom_list = new_bom.groupby(['Bill Number']).size().reset_index(name='Count')
new_bom_list = list(new_bom_list.itertuples(index=False, name=None))

### Define button locations
bill_num_pos = [1821, 452]
crate_bill_ok_btn_pos = [1382, 809]
additional_tab_pos = [619, 503]
ok_btn_pos = [1237, 1150]
line_tab_pos = [1677, 561]
finish_ok_btn_pos = [2083, 1195]

# Uncomment below to run each step
# get_cursor_position()
# time.sleep(2)
# click_on_screen(additional_tab_pos[0], additional_tab_pos[1])

### Run the script
item_code_loop(new_bom_list, new_bom)
