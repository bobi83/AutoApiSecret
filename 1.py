# -*- coding: UTF-8 -*-
import requests as req
import json,sys,time
#先注册azure应用,确保应用有以下权限:
#files:	Files.Read.All、Files.ReadWrite.All、Sites.Read.All、Sites.ReadWrite.All
#user:	User.Read.All、User.ReadWrite.All、Directory.Read.All、Directory.ReadWrite.All
#mail:  Mail.Read、Mail.ReadWrite、MailboxSettings.Read、MailboxSettings.ReadWrite
#注册后一定要再点代表xxx授予管理员同意,否则outlook api无法调用






path=sys.path[0]+r'/1.txt'
num1 = 0

def gettoken(refresh_token):
    headers={'Content-Type':'application/x-www-form-urlencoded'
            }
    data={'grant_type': 'refresh_token',
          'refresh_token': OAQABAAAAAAAm-06blBE1TpVMil8KPQ41EXN5InExgKtrKG3s8XA1JuCqKn45DrML_uWKRLeAvzVGxegL4rMDzFlEEWywrlm3uG3UKrMOvx0GGeYzxTRUWYuYaUba3Z9LxDf08vjYsYow3jCHlQ5jn16vqL4rOVYbIKO7qoqDVpquOv8B6P46D8LuL3NCmiAJx1u8U61_DCEQGZzXyuRX2YPWvagYFCddzJalc9aSGEJzTkDWrM3oeyMsRzc7ggsav578YdZAsCjw_htBsWerwgCe_VGtZxgriI1c8GtYRX0SfBCu5VGI6WteDDJfCPZcGCd49PV6kawieWyLPZ0GZRG43opioKWmT_iJ4Xlf12of2YqHsNEmPn0Dad8y1v4-PZ0E2NEdthrKPc032psSzeZxjan-y2c3AfWCL_neDcmNI6mwSUi65uA2zZmEM3wE6puSYgS33n9VwhLPIhrVb6vNzTtmpZYUR5AtegQacqtSnP_UcIcJ02goXOrmmimAGGdw2jfXs5W5VlrAFIQoSM1NlnTAlsIPe7dE6wBFCKJdrDdy7BdpqAI7hZuaqfhvxwJbYxkj7UNsx8130MCoC-aF8BHRPzwqvTq6r5J_e5qgNr-aDadGdV0kf9ePInSvBmqvhQXacQSEZ0DC2QktK9tOeY-qX6RrNEEMR-hEu1UecSaNsW2ilYYTpdI2b3q6gXUHgVB69bFvRzdXfyrxmeP8Xm4yz4Y5UyeAltKrvgX2lOszf3TY8E2eZcXpyTbfbkU8dGA01YR600yeO47xmUgkfy4b9S6CelJHBhVgpDaPRjgIAPaMqwaFR3jtBbafJcKDggdi0G_k3dEggtfmslxTsk_4HO2cZgk3kbUzKCJrAhuRGSeYsX4qorWqCUz0HlY8oe5l_2DffWpvndbnvW_DXnuUAlTszYSFBYKb5nBWOXXMMwFkoHY_CPH5_b22uzASJOc4RFq8ubIqbROcRlrVplC_5dY-on4AvisAfgiSgHHnj6_AbVbIx4FA7KyeJOtwoXB6h8oTjVDcFxiafl7dIWyXmDGRNHFFMqMZqktoPb_dfbhgOCAA,
          'client_id':cbe0813b-600b-4ecc-8f78-943241d92386,
          'client_secret':bp31r2asjg8G~7J3Ax.3Afa~0F6y6qTjZ~,
          'redirect_uri':'http://localhost:53682/'
         }
    html = req.post('https://login.microsoftonline.com/common/oauth2/v2.0/token',data=data,headers=headers)
    jsontxt = json.loads(html.text)
    refresh_token = jsontxt['refresh_token']
    access_token = jsontxt['access_token']
    with open(path, 'w+') as f:
        f.write(refresh_token)
    return access_token
def main():
    fo = open(path, "r+")
    refresh_token = fo.read()
    fo.close()
    global num1
    localtime = time.asctime( time.localtime(time.time()) )
    access_token=gettoken(refresh_token)
    headers={
    'Authorization':access_token,
    'Content-Type':'application/json'
    }
    try:
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root',headers=headers).status_code == 200:
            num1+=1
            print("1调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive',headers=headers).status_code == 200:
            num1+=1
            print("2调用成功"+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/drive/root',headers=headers).status_code == 200:
            num1+=1
            print('3调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/users ',headers=headers).status_code == 200:
            num1+=1
            print('4调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/messages',headers=headers).status_code == 200:
            num1+=1
            print('5调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messageRules',headers=headers).status_code == 200:
            num1+=1
            print('6调用成功'+str(num1)+'次')    
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders/Inbox/messages/delta',headers=headers).status_code == 200:
            num1+=1
            print('7调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/drive/root/children',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次')
        if req.get(r'https://api.powerbi.com/v1.0/myorg/apps',headers=headers).status_code == 200:
            num1+=1
            print('8调用成功'+str(num1)+'次') 
        if req.get(r'https://graph.microsoft.com/v1.0/me/mailFolders',headers=headers).status_code == 200:
            num1+=1
            print('9调用成功'+str(num1)+'次')
        if req.get(r'https://graph.microsoft.com/v1.0/me/outlook/masterCategories',headers=headers).status_code == 200:
            num1+=1
            print('10调用成功'+str(num1)+'次')
            print('此次运行结束时间为 :', localtime)
    except:
        print("pass")
        pass
for _ in range(3):
    main()
