l = [1,2,3,4,2,3,45,23,2,3,4]
c = max(set(l), key=l.count)
print(c)

f = filter(lambda x : x > 2, l)
print(list(f))

s = "".join([f"{i}" if i > 0 else "NULL" for i in range(10)])
#print(s)