#!/usr/bin/env python
# coding: utf-8

# In[1]:


students_list = ["Namya","Heet","Yug"]
print("Original list -", students_list )
students_list.append("Viana")
students_list[1] = "Adeetya"
students_list.remove("Yug")           
print("New List -",students_list )
students_tuple = ("Namya", "Heet","Yug")
print("Original Tuple-",students_tuple )
students_tuple = students_tuple + ("Viana",)
change_tuple = list(students_tuple)
change_tuple[1] = "Adeetya"
students_tuple = tuple(change_tuple)
temp_tuple = list(students_tuple)
temp_tuple.remove("Yug")
students_tuple = tuple(temp_tuple)
print("New Tuple-",students_tuple)
students_dict = {1:"Namya",2:"Heet",3:"Yug"}
print("Original Dictionary-",students_dict)
students_dict[4] = "Viana"
students_dict[2] = "Adeetya"
del students_dict[3]
print("New Dictionary-",students_dict)


# In[ ]:




