# +
*** Settings ***
Library          RPA.Desktop.Windows
Library          RPA.Desktop



*** Tasks ***
Open Outlook using run dialog
  
    ${app1}  Open Using Run Dialog   outlook   Inbox   wildcard=True   timeout=30
    ${position}=    Get mouse position
    Log    Current mouse position is ${position.x}, ${position.y}
    Sleep  10
    Move Mouse  point:700,300
    Sleep  5
    Press Keys  ctrl  n
    Sleep  5
    Type text  ramkumar.r@blueconchtech.com
    Sleep  5
    Press Keys  enter
    Sleep  5
    Press Keys  alt  u
    Sleep  5
    Type text  Test Subject
    Sleep  5
    Move Mouse  point:657,103
    Sleep  5
    Mouse Click  method=coordinates  x=657   y=103
    Sleep  5
    Move Mouse  point:743,896
    Sleep  5
    Mouse Click  method=coordinates  x=743   y=896
    Sleep  5
    Type text  aadhar.jpg
    Sleep  5
    Press Keys  enter
    Sleep  5
    Move Mouse  point:28,256
    Sleep  5
    Mouse Click  method=coordinates  x=28  y=256
    Sleep  5

    


*** Tasks ***
Minimal task
    Log  Done.


    
        



# -



outlook







