import json
import os
from datetime import datetime
filename= "committee_data.json"
member_list={}
total_collection = 0

def save_data():
    with open(filename,"w") as file:
        json.dump(member_list,file)
        print("Data stored successfully")


def load_data():
    global member_list,total_collection
    if os.path.exists(filename):
        with open(filename,"r") as file:
            data=json.load(file)
            member_list = {int(k): v for k, v in data.items()}
        temp_total=0
        done_members=0
        for m_id in member_list:
            temp_total+=member_list[m_id]["balance"]  
            if member_list[m_id]["committee_status"]=="Done":
                done_members+=1
        payout_amount = done_members * (len(member_list) * 5000)
        total_collection = temp_total - payout_amount
        print("Previous data loaded successfully!")
    else:
        member_list = {}
        total_collection = 0
        print("No previous record found. Starting fresh.")
        
        

def register_new_member():
    print("\n" + "-"*40)
    print("register new member")
    
    name=(input("please enter a new member name")).title().strip()
    id=(input("enter a member id no"))
    gurantor_name=input("please enter the name of your gurantor").title().strip()
    gurantor_id=(input("enter gurantor id no "))
    if name.replace(" ", "").isalpha() and gurantor_name.replace(" ", "") .isalpha() and id.isdigit() and gurantor_id.isdigit():
        id=int(id)
        gurantor_id=int(gurantor_id)
        if id!=gurantor_id:
        

            if id not in member_list:
                if not member_list or gurantor_id in member_list:

                    


                    member_data={
                        'Member_Name':name,
                        'Member_id':id,
                        'Member_Gurantor':gurantor_name,
                        "gurantor_id":gurantor_id,
                        "score":100,
                        "is_paid":False,
                        "balance":0,
                        "history": {},
                        "committee_status":"Not Done"
                        
                    }
                    member_list[id]=(member_data)
                    save_data()
                    print(f"\nSUCCESS: {name} has been registered!")

                
                else:
                    print("you need to choose gurantor from already registerd members")     
            else:
             print("ID already exist, please choose a unique one!")
        else:
            print(" gurantor and user id should be different")   
    else:
        print("member name and guarantor name should be in alphabets\n member id and gurantor id should be in didgit")  
    print("\n" + "-"*40)              
    
def show_registered_members():
    print("\n" + "-"*40)
    if not member_list:
        print("No member registered yet")
    else:
     for key,value in member_list.items():
         print(key,":",value)
    print("\n" + "-"*40)     
def record_payment():
    print("\n" + "-"*40)
    print("Record payments")
    search_id=(input("enter a member id "))
    if search_id.isdigit():
         search_id=int(search_id)
         global total_collection
        
            
         if search_id in member_list:
            
            amount=(input("enter a amount"))
            date=(input("enter a date"))
            month_number=(input("enter a month number (1 - 12)"))
            if amount .isdigit() and date.isdigit() and month_number.isdigit():
                amount=int(amount)
                date=int(date)
                month_number=int(month_number)
                

            
                
                if month_number >= 1 and month_number <=12 and date>=1 and date <=31 :
                    months_map = ["Jan", "Feb", "March", "April", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"]
                    month_name = months_map[month_number - 1]
                    if month_name not in member_list[search_id].get("history",{}):

                     if amount == 5000:

                        if date > 5:
                            p_type='late'
                            member_list[search_id]["score"]-=10
                            g_id=member_list[search_id]["gurantor_id"]
                            
                            if g_id in member_list:
                                member_list[g_id]["score"]-=3
                                
                        else:
                            p_type ="ontime"
                            member_list[search_id]["score"]+=5


                        member_list[search_id]["balance"]+=amount
                        member_list[search_id]["history"][month_name] = {}
                        member_list[search_id]["history"][month_name]["payment_type"] = p_type
                        member_list[search_id]["is_paid"]=True


                        
                        total_collection += amount

                        target = len(member_list)*5000

                        if total_collection >= target:
                            winner_id = None
                            max_score = -1

                            for m_id in member_list:
                                m_data = member_list[m_id]
                                if m_data["score"] > max_score and m_data.get("committee_status") != "Done":
                                    max_score = m_data["score"]
                                    winner_id = m_id
                            
                            if winner_id is not None:
                                    member_list[winner_id]["committee_status"] = "Done"
                                    member_list[winner_id]["payout_date"] = datetime.now().strftime("%d-%b-%Y")
                                    
                                    
                                    total_collection -= target 
                                    
                                    print(f"\nSUCCESS: Committee paid to {member_list[winner_id]['Member_Name']}!")
                                    print(f"Remaining balance in collection: Rs. {total_collection}")
                            else:
                                print("Payment recorded for", month_name , "successfully!")
                            

                    
                    
                        save_data()  

                     else:
                        print("Please enter the exact amount of the committee")
                    else:
                        print("your payment is already done")     
                else:
                    print("Enter a valid month number (1 - 12) and valid date (1-31)")
            else:
                print("amount, date and month number should be in digit")        

         else:
          print("member not found ")
    else:
       print("Search id should be in didgit") 
    print("\n" + "-"*40)
def delete_payment_record():
    print("\n" + "-"*40)
    print("delete payments")
    global total_collection
    search_id = (input("Enter member id which you want to delete  "))
    if search_id .isdigit():
        search_id=int(search_id)
    
        if search_id in member_list:
            if "history" in member_list[search_id] and member_list[search_id]["history"]:
            
                month_to_del = int(input("enter a month number which you want to del "))
                
                if month_to_del >= 1 and month_to_del <=12:

                    months_map = ["Jan", "Feb", "March", "April", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"]
                    month_to_del = months_map[month_to_del - 1]
                    if month_to_del in member_list[search_id]["history"]:
                        p_status = member_list[search_id]["history"][month_to_del]["payment_type"]
                        if p_status == "late":
                        
                            member_list[search_id]["score"] += 10
                            g_id = member_list[search_id]["gurantor_id"]
                            if g_id in member_list:
                                member_list[g_id]["score"] += 3
                        elif p_status == "ontime":
                            
                            member_list[search_id]["score"] -= 5        
                    
                        del member_list[search_id]["history"][month_to_del]
                        
                    
                        member_list[search_id]["balance"] -= 5000
                        total_collection -= 5000

                        if not member_list[search_id]["history"]:
                            member_list[search_id]["is_paid"] = False
                        
                        save_data()
                        print(f"{month_to_del} entry is deleted succesfully ")
                    else:
                        print("There is no entry of the",month_to_del )
                else:
                    print("Please enter a number bwtween 1 and 12")
            else:
                print(" THIS ENTRY IS NOT RECORDED ")
        else:
            print("Member ID NOT FOUND ")
    else:
        print(" search id should be in digit ") 
    print("\n" + "-"*40)           

def reset_system_for_new_committee():
    print("\n" + "-"*40)
    confirm=input("enter confirmation yes or no ").upper()
    if confirm.isalpha()and confirm== "YES":
        global member_list
        member_list = {}
        with open(filename, "w") as file:
            json.dump(member_list, file)
        print("Data reset succesfully")  
    else:
        print("by not entering yes , data can not be reset and previous data is secured") 
    print("\n" + "-"*40)                 


def check_pending_alerts():
    print("\n" + "-"*40)
    print("      PENDING PAYMENTS & RISK AUDIT")
    print("-"*40)

    current_month_num = datetime.now().month 
    months_map = ["Jan", "Feb", "March", "April", "May", "June", "July", "Aug", "Sept", "Oct", "Nov", "Dec"]

    to_be_removed = []

    if not member_list:
        print("No members to audit")

    for m_id, d in member_list.items():
        paid_months = d.get("history",{}).keys()
        pending_count = 0

        for i in range(current_month_num):
            m_name = months_map[i]
            if m_name not in paid_months:
                pending_count += 1

        print(f"\nMember: {d['Member_Name']} (ID: {m_id})")
        print(f"Total Pending Months: {pending_count}")

        if 1 <= pending_count <= 2:
            print(f"Dear Member, Your {pending_count} months payment is pending")
        elif 3 <= pending_count <= 4:
            print(f"Warning Guarantor ({d['Member_Gurantor']}): \n If he dont pay till next month, the payment will be deducted from your account!")
        elif pending_count >= 5:
            refund_amount = d['balance'] * 0.5
            print(f"{d['Member_Name']} is removed from the committee. \n 50% amount is deducted as penalty because you have pending payments for more 5 months. \n Your remaining amount is: {refund_amount}")
            to_be_removed.append(m_id)

    for m_id in to_be_removed:
        del member_list[m_id]

    if to_be_removed:
        save_data()
    
    print("\n" + "*"*40)

def show_summary_report():
    global total_collection
    print("\n" + "="*35)
    print("      MONTHLY SUMMARY REPORT")
    print("="*35)
    
    pending_members = 0
    
    for m_id, d in member_list.items():
        
        if not d['is_paid']:
            pending_members += 1
    
    print(f"Total Money Collected: Rs. {total_collection}")
    print(f"Members Pending: {pending_members}")
    print("="*35)

load_data()
while True:
    choice=(input("""enter a number
                     {1} for registering new member 
                     {2} for showing all the registered members 
                     {3} for record payments
                     {4} for delete record 
                     {5} for summary report 
                     {6}check pending payments
                     {7} for reset the previous data and start new committee
                     {8} for exit the program  """))
    if choice.isdigit():
        choice=int(choice)
        if choice==1:
            register_new_member()
        elif choice==2:
            show_registered_members()   
        elif choice==8:
            print("thanks for using , system closed succesfully") 
            save_data()    
            break
        elif choice==3:
            record_payment() 
        elif choice==5:
            show_summary_report()
        elif choice==4:
            delete_payment_record()
        elif choice==6:
            check_pending_alerts() 
        elif choice==7:
            reset_system_for_new_committee()                 
        else:
            print("invalid choice ") 
    else:
        print("choice should be in digit")        