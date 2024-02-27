filename = "graficos.ini"

i = 1
with open(filename, 'r', encoding='windows-1252') as file:
    while (line := file.readline()):
        if line[0:3] == "Grh":
            index = line.split("=")[0][3:]
            if index != str(i):
                print("Void", i)
                i = int(index) + 1
            else:
                i+=1