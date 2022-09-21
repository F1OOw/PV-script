import aspose.words as aw

class Pv_infos:
    def __init__(self,name,theme,date,hour,place,presenter,presents,points):
        self.name = name
        self.theme = theme
        self.date = date
        self.hour = hour
        self.place=place
        self.presenters = presenter
        self.presents = presents
        self.points = points
    def display_infos(self):
        print(f'PV file name: {self.name}')
        print(f'Meeting theme: {self.theme}')
        print(f'Meeting date: {self.date}')
        print(f'Meeting hour: {self.hour}')
        print(f'Meeting place : {self.place}')
        print(f'Meeting presenter/s: {", ".join(self.presenters)}')
        print("Present members:")
        for i in self.presents:
            print(f"     >{i}")


# Main menu 
def main_menu(): 
    print("-----------------Menu-----------------")
    print("     1- Start a new PV.")
    print("     2- Save")
    print("     3- Exit")
    choice = int(input("Your choice: "))
    print("--------------------------------------")
    if choice not in range(1,4):
        print("Not valid choice, try again")
        main_menu()
    else:
        return choice


# Present members in the metting
def get_presence():
    presents_list = []
    print("Enter the names of present members: ")
    print("When you're done enter the keyword 'end'")
    while (True):
        print("     >",end="")
        name = input()
        if (name=="end"):
            break
        elif (name== ""): #  add some criterias here
            continue
        else:
            presents_list.append(name)
    return presents_list

def fetch_pv_infos():
    name = input("PV file name: ")
    name = name.split(".")[0]
    theme = input("Enter the theme: ")
    date = input("Enter a date: ")
    hour = input("Enter the hour of the meeting: ")
    place = input("Enter the place/platform of the meeting: ")
    presenter = input("Who is/are presenting (separe by commas): ")
    presenter = [i.strip() for i in presenter.split(",")]
    present_members = get_presence()
    points = []
    pv_infos = Pv_infos(name,theme,date,hour,place,presenter,present_members,points)
    return pv_infos

def pv_menu():
    print("1- Add a Discussed point")
    print("2- Add notes to a Discussed point")
    print("3- Quit")
    choice = int(input("Your choice: "))
    if (choice not in range(1,4)):
        print("Invalid choice, Try again")
        pv_menu()
    else:
        return choice
def add_point(pv):
    title = input("Title: ")
    if (title == ""):
        print("Try again")
        add_point(pv)
    else:
        notes = []
        print("Add your notes here:")
        print("Stop keyword = 'end'")
        while (True):
            print("     >", end ="")
            note = input()
            if (note == ""):
                continue
            elif(note == "end"):
                break
            else:
                notes.append(note)
        pv.points.append([title,notes])
            
def add_notes(pv):
    if (len(pv.points)==0):
        print("No titles available yet")
    else: 
        print("Available titles:")
        for i in pv.points:
            print(i[0])
        print("--------------------------------------")
    
        print("Please chose the point title: ", end ="")
        title = input()
        found = False
        for i in pv.points:
            if (i[0] == title):
                found = True
                appended_notes = []
                print("Add your notes and stop with 'end': ")
                while (True):
                    print("     >",end="")
                    note = input()
                    if (note==""):
                        continue
                    elif (note == "end"):
                        break
                    else:
                        appended_notes.append(note)
                i[1] += appended_notes
        if (not found):
            print("The title name you entered is not correct")

def new_pv():
    pv = fetch_pv_infos()
    print("These are your PV informations: ")
    pv.display_infos()
    print("-----------------------------------------------")
    while(True):
        choice = pv_menu()
        if (choice == 1):
            add_point(pv)
        elif (choice == 2):
            add_notes(pv)
        elif (choice == 3):
            break
    return pv

def save(pv):
    f = open(pv.name+".txt", "w")
    f.write("Theme: {}\nDate: {}\nHour: {}\nPlace: {}\nPresenter/s: {}\nPresent members:\n-{}\n".format(pv.theme,pv.date,pv.hour,pv.place,", ".join(pv.presenters), "\n-".join(pv.presents)))
    f.write("Discussed points: \n")
    for i in pv.points:
        f.write(i[0]+":\n")
        for j in i[1]:
            f.write("   >{}\n".format(j))
    f.close()
    #Transform to word
    doc = aw.Document(pv.name+".txt")
    doc.save(pv.name+".docx")


while(True):
    choice = main_menu()
    if (choice==1):
        All_info = new_pv()
    elif (choice == 2):
        save(All_info)
    elif (choice == 3):
        print("Do you want to save the changes first ? [Y/N]: ",end="")
        if (input().upper()=="Y"):
            save(All_info)
        break


