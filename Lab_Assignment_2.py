#!/usr/bin/env python
# coding: utf-8

# In[1]:


a = int(input("Enter first number "))
b = int(input("Enter second number "))
c = int(input("Enter third number "))
if a == b:
          if b == c:
                    print("All numbers are equal")
          elif b > c:
                     print(" a and b are larger ")
          else:
                print("c is larger")
elif b ==c:
              if b > a:
                     print(" b and c are larger")
              else:
                   print("a is larger")
else:
     print("b is larger ")


# In[ ]:




