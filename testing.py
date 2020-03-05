
from win32com.client import Dispatch         ##used for voice output
speak=Dispatch("SAPI.SpVoice")

from abc import ABC, abstractmethod


import pickle           ##it is used to serialize and deserialized a python object
import time             ##used to delay final results
import projectmodules
from projectmodules import module1
from projectmodules import module2
from projectmodules import module3



class AccountError(Exception):
    pass

class AccountTypeError(Exception):
    pass

class NotFoundError(Exception):
    pass


class AmountLargeError(Exception):
    pass

class ReturnError(Exception):
    pass

class DominationError(Exception):
    pass

class CorrectError(Exception):
    pass


class AmountisGreaterError(Exception):
    pass

class AmountisEqualError(Exception):
    pass




class CreateAccount():
    """ This class is creating a new account in bank. """
    def creating(self):

        speak.Speak("welcome to the create account class")
        speak.Speak("here you would be creating your account in our bank")
        time.sleep(2.0)
        
        
        l=[]
        D={}
        
        file=open("new_one.pickle","rb")

        l=pickle.load(file)            #this is deserialization
        
        
        
        
        self.accountno=int(input("Enter the account no:"))
        self.atm=int(input("Enter ATM pin code:"))
        self.cnic=int(input("Enter CNIC number:"))
        try:
            for accounts in l:
                if (accounts["accountno"]== self.accountno):
                    raise AccountError
                else:
                    continue
        except:
            time.sleep(2.4)
            print("oops, looks like this account already exists!")
        else:
            self.name=input("Enter the name:")
            self.type=input("Enter the type [current/saving]:")
            try:
                if self.type != "current" and self.type != "saving":
                    raise AccountTypeError

            except:
                time.sleep(2.4)
                print("Please enter a valid account type!")
            else:
                self.amount=int(input("The initial amount is Rs"))
        
        
                D["accountno"] = self.accountno
                D["name"] = self.name
                D["type"] = self.type
                D["amount"] = self.amount
                D["ATM pin code"]=self.atm
                D["CNIC number"]=self.cnic
        
                l.append(D)
                

            
                file=open("new_one.pickle","wb")
                pickle.dump(l,file)             #this is serialization
                file.close()
                
                time.sleep(2.4)           
                print("Your account has been created!")
                speak.Speak("your account has been created")
    
        
        

        



       
class ShowAccount():              ## this class is the container class     #composition
    """ Here you can view your account and can easily deposit and withdraw money from it. """
    
    def show_account(self,account_no,atm_no):
        
        
        file=open("new_one.pickle","rb")
        l=pickle.load(file)
##        print(l)
       
        accountFound=False
        for account in l:
            if(account["accountno"]== account_no and account["ATM pin code"]==atm_no):
                accountFound=True
                time.sleep(2.4)
                print("Account holder's name is:", account["name"])
                print("Account type is:", account["type"])
                print("Your current balance is Rs",account["amount"])


                try:
                    speak.Speak("Do you want to deposit or withdraw amount?:")
                    wht=input("Do you want to withdraw amount/deposit amount?:")

                    if wht=="withdraw":
                        obj4=module2.Withdraw()
                        obj4.withdraw(key)
                    elif wht=="deposit":
                        obj3=module2.Deposit()
                        obj3.deposit(key)
                    elif wht!="withdraw" and wht!="deposit":
                        raise CorrectError

                except CorrectError as anerror:
                    time.sleep(2.4) 
                    print("Please choose from the given options!")
                    
                    

                
   

            else:
                continue
           
        try:
            if accountFound==False:
                raise NotFoundError
        except:
            time.sleep(2.4)

            print("oops,looks like this account doesn't exist with this ATM pin code!")
            speak.Speak("looks like this account doesn't exist with this ATM pin code!")


 

class ChangeAccountNo():
    """ This class is created so that a user can change his/her account number on providing some basic information."""
    def change_account(self,account_no,atm_no):
        file=open("new_one.pickle","rb")
        l=pickle.load(file)
        new_accountno=int(input("Enter a new account no:"))
        accountFound=False
        for y in l:
            if( y["accountno"]==account_no and y["ATM pin code"]==atm_no ):
                
                
                print("Your previous account number was: ",y["accountno"])
                y["accountno"]=new_accountno
                time.sleep(2.4) 
                print("Your present account number is :",y["accountno"])
                file=open("new_one.pickle","wb")
                pickle.dump(l,file)
                file.close()
                accountFound=True
                time.sleep(2.0)
                print("Your account number has been changed!")
                speak.Speak("Your account number has been changed")
        if accountFound==False:
           time.sleep(2.4) 
           print("Account not found with this ATM pin code . Account number couldnot be changed")
           speak.Speak("Account not found with this ATM pin code .   Account number couldnot be changed")

           

class ChangeAccountName():
    """ This class is created to change the user's current account name on providing some basic information."""

    def change_name(self,account_no,atm_no):
       
         file=open("new_one.pickle","rb")
         l=pickle.load(file)
         new_name=input("enter new name of account")
         accountFound=False
         for name in l:
             if(name["accountno"]==account_no and name["ATM pin code"]==atm_no ) :
                 
                 print("Your previous account name was",name["name"])
                 name["name"]=new_name
                 time.sleep(2.4) 
                 print("Your new account name is",name["name"])
                 file=open("new_one.pickle","wb")
                 pickle.dump(l,file)
                 file.close()
                 accountFound=True
                 time.sleep(2.0)
                 print("Your account name has been changed")
                 speak.Speak("Your account name has been changed")
         if accountFound==False:
            time.sleep(2.4) 
            print("Account not found with this ATM pin code . Account name couldnot be changed")
            speak.Speak(" sorry but the account not found with this ATM pin code .   Account name couldnot be changed")






class MoneyTransfer():
    """ This class deals with the transactions made in the bank ."""
    def transfer(self,account_no,cnic_no):

        l=[]
        D={}
        file=open("new_one.pickle","rb")
        l=pickle.load(file)
        

        
        accountFound=False
        for account in l:
            if (account["accountno"]==account_no and account["CNIC number"]==cnic_no):
                accountFound=True
                self.name=input("enter your full name:")
                self.receiver=input("enter receiver's name:")
                self.accountno2=int(input("enter receiver's account number"))
                self.city=input("enter the city")
                self.amount2=int(input("the amount to be transferred is Rs"))
                for x in l:
                    if (x["accountno"]==account_no and account["CNIC number"]==cnic_no):
                        time.sleep(2.4) 
                        print("your current balance is Rs",x["amount"])
                        try:
                            if (x["amount"] < self.amount2):
                                raise AmountisGreaterError
                            elif (x["amount"] == self.amount2):
                                raise AmountisEqualError

                        
                        except AmountisGreaterError as error1:
                            time.sleep(2.4) 
                               
                            print("Sorry,the amount you entered is more than your current balance!")
                            speak.Speak("Sorry,  but the amount you entered is more than your current balance!")
                    
                        except AmountisEqualError as error2:
                            x["amount"]-=self.amount2
                            file=open("new_one.pickle","wb")
                            pickle.dump(l,file)
                            file.close()
                            time.sleep(2.4)  
                    
                            print("Your new balance is Rs 0.00")
                            speak.Speak("your account is now empty")
                    

                        else:
                            x["amount"]-=self.amount2
                            file=open("new_one.pickle","wb")
                            pickle.dump(l,file)
                            file.close()
                            print("money transfer was successful!")
                            speak.Speak("money transfer was successful!")
                            time.sleep(2.4)
                            speak.Speak("this is your new balance")
                            print("Now your new balance is Rs",x["amount"])
                            time.sleep(2.4)
                            speak.Speak("hope you are satisfied by the transaction")
            else:
                continue

        if accountFound==False:
            time.sleep(2.4) 
            print("Sorry , this account doesn't match with this CNIC number!")
            speak.Speak("sorry but your c n i c number does not match")

class PayementOfBills:
    """This class is created so that a user can clear his bills on time."""

    def __init__(self,date):

        
        if  type(date) is str:
            self.date=int(date[0:2])
            
        else:
            self.date=date

    def bills(self):

        self.month=input("Enter the month of payement:")
        self.amount=int(input("Your bill is Rs"))
        self.address=input("Enter your house address:")

    def __gt__(self,other):
        """This function compares the date on which the bill is payed with the due date of bill payement"""
        if (self.date>other.date):
            time.sleep(2.4) 
            print("Sorry, but you would have to pay Rs",self.amount+50,"since the due date has passed!")
            speak.Speak("sorry but you'll have to pay this additional amount because of late payement of your bill")
        else:
            time.sleep(2.4) 
            print("Your bill of Rs",self.amount,"for the month of",self.month,"has been cleared! ")
            speak.Speak("your bill has been cleared")
            time.sleep(2.0)
            speak.Speak("hope you are satisfied")
            


     
class CustomerService:
    """The purpose of this class is that, whenever a person enters a bank he should be entertained on the basis of his token number in comparison with other customers."""
    
    def __init__(self,token):

        self.token=token
        
    def __lt__(self,other):
        speak.Speak("welcome to the customer service class")
        speak.Speak("here we would access you as per your token number")
        time.sleep(2.0)
        
        if(self.token<other.token):
            
            return True

        

        else:
            
            return False           




        
    
    


obj1=CreateAccount()
obj2=ShowAccount()

obj5=ChangeAccountNo()
obj6=ChangeAccountName()

obj7=module3.Business_loans()
obj8=module3.House_loans()
obj9=module3.Car_loans()
obj10=MoneyTransfer()
obj11=module1.PrizebondNumbers()
obj12=module1.PrizebondAward()


print("*********WELCOME TO THE BANKING HIERARCHY*********")
speak.Speak("welcome to the banking Hierarchy")
####speak.Speak("This project is inspired by the idea of incorporating modern banking system with the IT world.   The purpose of this project is to create a user-friendly banking system,  in which all banking actions are performed.")


answer="yes"                          
x = "   "
while answer =="yes":
    
    speak.Speak("Listed below are some available operations that you can perform using this system")
    print("******MAIN MENU******")
    print("\t1- Create a new account")
    print("\t2- Deposit or Withdraw")
    print(" \t3- Change account number")
    print("\t4- Change account name")
    print("\t5- Take loans for business")
    print("\t6- Take loans for house")
    print("\t7- Take loans for car")
    print("\t8- Money Transaction")
    print("\t9- Prize bond Input")
    print("\t10- Prize bond Award")
    print("\t11- Payement of Bills")
    print("\t12- Customer Dealings")

    print("\t13- Exit")
    speak.Speak("please select your choice")
    print("select your choice [1-13]")
    x = input()
    speak.Speak("you'll be getting your results shortly")

    if x == "1":
        obj1.creating()
    elif x == "2":
        key = int(input("enter the account no:"))
        k=int(input("enter your ATM pin code"))

        obj2.show_account(key,k)

    elif x == "3":
        key = int(input("enter the account no:"))
        k=int(input("enter your ATM pin code"))
        obj5.change_account(key,k)
    elif x == "4":
        key = int(input("enter the account no:"))
        k=int(input("enter your ATM pin code"))
        obj6.change_name(key,k)
    elif x == "5":
        
        obj7.input()
        obj7.abstract()
    elif x=="6":
        obj8.input()
        obj8.abstract()
        
        
    elif x=="7":

        obj9.abstract()
    elif x=="8":
        speak.Speak("welcome to the money transfer class")
        speak.Speak("here you can transfer money from your account to another account")
        key=int(input("enter your account no"))
        k=int(input("enter your CNIC number"))
        obj10.transfer(key,k)

    elif x == "9":    
        obj11.new1()
        obj11.new2()
        
    elif x=="10":

        obj12.awards(obj11)

    elif x=="11":
        obj13=PayementOfBills((input("enter today's date :")))
        obj14=PayementOfBills(25)
        obj13.bills()
        obj15=obj13>obj14

    elif x=="12":

        obj16=CustomerService(int(input("Token no of first customer is:")))
        obj17=CustomerService(int(input("Token no of second customer is:")))
        if(obj16<obj17):
            time.sleep(2.4) 
            print("customer one should be dealed first")
            speak.Speak("customer one should be dealed first")
            
        
        else:
            time.sleep(2.4) 
            print("customer two should be dealed first")
            speak.Speak("customer two should be dealed first")


    else:
        break

    
    time.sleep(2.0)
    speak.Speak("Do you want to continue")
    answer=input("Do you want to continue?[yes/no] :")
    
    if answer=="yes":
        time.sleep(2.4)  
        print("******Welcome To The Main Menu******")
        speak.Speak("you are again welcome to the main menu")
        continue

    elif answer=="no":
        speak.Speak("hope you were satisfied by your results")
        speak.Speak("thanks for using the banking hierarchy")
        break

    elif answer!="yes" and answer!="no":
        time.sleep(2.4) 
        print("this is not a valid answer")
        break

    






        
    
    

