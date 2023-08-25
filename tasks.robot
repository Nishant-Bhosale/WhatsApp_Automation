*** Settings ***
Documentation       Playwright template.

Library             RPA.Browser.Selenium
Library             RPA.Excel.Files
Library             RPA.Tables
Library             Collections
Library             RPA.FTP
Library             RPA.Desktop
Library             RPA.FileSystem
Library             RPA.Browser.Playwright
Library             OperatingSystem
# google-chrome --remote-debugging-port=9289 --user-data-dir="/home/vijender/.config/google-chrome/Default"


*** Variables ***
${sendingImage}                 ${False}
${EXCEL_FILE}                   /home/vijender/Desktop/shopout/whatsappRcc/contact.xlsx
${file_path}                    /home/vijender/Desktop/shopout/whatsappRcc/file.png
${message_path}                 /home/vijender/Desktop/shopout/whatsappRcc/message.txt
${new_chat_element}             //div[@title="New chat"]
${search_chat_element}          //div[@title="Search input textbox"]
${first_el_of_chat_search}      //div[@class="_199zF _3j691"]
${msg_input_el}                 //*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]
${image_upload_input_el}        //*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div/div/span/div/ul/div/div[1]/li/div/input
${image_caption_input_el}       //*[@id="app"]/div/div/div[3]/div[2]/span/div/span/div/div/div[2]/div/div[1]/div[3]/div/div/div[2]/div[1]/div[1]
${success_contacts_file}        /home/vijender/Desktop/shopout/whatsappRcc/success_contacts.txt
${failed_contacts_file}         /home/vijender/Desktop/shopout/whatsappRcc/failed_contacts.txt


*** Tasks ***
Open Browser
    Open Workbook    ${EXCEL_FILE}
    ${worksheet}=    Read Worksheet
    ${orders}=    Create Table    ${worksheet}
    Attach Chrome Browser    9289
    RPA.Browser.Selenium.Go To    https://web.whatsapp.com
    ${message}=    Read File    ${message_path}
    Set Clipboard Value    ${message}
    ${copied_message}=    Get Clipboard Value
    FOR    ${element}    IN    @{worksheet}
        ${values}=    Get Dictionary Values    ${element}

        # click new chat element
        Send A Message To contact    ${values[0]}    ${message}
    END


*** Keywords ***
Send A Message To contact
    [Arguments]    ${contactNum}    ${message}

    TRY
        Wait Until Page Contains Element    ${new_chat_element}    20s
        Click Element    ${new_chat_element}
        Input Text    ${search_chat_element}    ${contactNum}
        Sleep    1s
        Click Element    ${first_el_of_chat_search}
        Wait Until Page Contains Element    ${msg_input_el}    5s

        IF    ${sendingImage}
            Upload File    ${file_path}    ${message}
        ELSE
            Plain Text Message    ${message}
        END
        OperatingSystem.Append To File    ${success_contacts_file}    ${contactNum}\n
    EXCEPT
        RPA.Browser.Selenium.Press Keys    ${None}    CTRL+a    # enter key
        OperatingSystem.Append To File    ${failed_contacts_file}    ${contactNum}\n
    END

Upload File
    [Arguments]    ${file_path}    ${caption}
    Click Element    //div[@data-testid="conversation-clip"]
    Sleep    0.2s
    Choose File    ${image_upload_input_el}    ${file_path}    # upload file
    Sleep    0.5s
    RPA.Browser.Selenium.Press Keys    ${None}    CTRL+v    # enter key
    Sleep    0.2s
    RPA.Browser.Selenium.Press Keys    ${None}    \ue007    # enter key

Plain Text Message
    [Arguments]    ${message}

    # Input Text    ${msg_input_el}    ${message}
    RPA.Browser.Selenium.Press Keys    ${None}    CTRL+v    # enter key

    RPA.Browser.Selenium.Press Keys    ${None}    \ue007    # enter key
    RPA.Browser.Selenium.Press Keys    ${None}    \ue00c    # escape key