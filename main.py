import msvcrt

### fontend menu ###
def Frontend_menu():
    print()
    print()
    print('       __________________________________________________________________________________________________________________       ')

    print("       Welcome to 'TH for distribution' Accounting Assistant")
    print('       made by Alfred Tharwat')
    print('       __________________________________________________________________________________________________________________       ')
    print('       [1] Sales Assistant')
    print()
    print()
    print()
    print('       __________________________________________________________________________________________________________________       ')
    print('       Enter a menu option in the Keyboard [1,2,3,4] :')
    print()
    print()
### calling other programs ###
Frontend_menu()
#calling the Acconting Assistant


while True:
    if msvcrt.getch().decode('utf-8') ==  '1':
        print('Initializing Sales Asisstant.....')
        import Sales_Assistant
        print('Task completed successfully ')
        break
while True:
        user_input = input("Enter 'exit' to quit: ")
        if user_input == 'exit':
            print("Exiting the program.")
            break