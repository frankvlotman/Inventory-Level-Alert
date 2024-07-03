import win32com.client as win32

# Dummy inventory data
inventory = {
    'item1': 10,
    'item2': 5,
    'item3': 0,
    'item4': 0,
    'item5': 3,
}

def send_email(items_to_order):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = 'example@gmail.com'
    mail.Subject = "Stock Reorder Notification"
    mail.Body = f"The stock level of the following items has reached 0. Please reorder:\n\n"
    for item in items_to_order:
        mail.Body += f"- {item}\n"
    mail.Send()

def check_inventory():
    items_to_order = []
    for item, stock_level in inventory.items():
        if stock_level == 0:  # Can edit this to a different value such as < less than a certain value
            items_to_order.append(item)
    if items_to_order:
        send_email(items_to_order)
        print("Email notification sent for items to reorder.")
    else:
        print("All items are in stock.")

if __name__ == "__main__":
    check_inventory()
