for i in range(10):
    print(i)

s = "".join([f"{i}" if i > 0 else "NULL" for i in range(10)])
print(s)