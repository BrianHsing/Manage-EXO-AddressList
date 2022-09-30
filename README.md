# 使用 Powershell 搭配 EXO v3 管理 Exchange Online 通訊清單

在企業內部有佈署 Exchange Server 的管理者，為了讓使用者能夠依據部門定義階層，方便企業人員尋找通訊錄中的信箱使用者，通常都會在 ECP 中透過通訊清單的功能去設定，但在移轉至 Exchange Online 後，發現無法使用網頁介面去設定，沒想到 Exchange Online 發展了這麼多年，這個功能依然還沒加到系統管理中心，還是只能使用 PowerShell 連線至 Exchange Online 進行設定。此篇就是教您如何透過 PowerShell 搭配 EXO V3 module 來建立通訊清單，Exchange Online 在此篇會以 EXO 縮寫表示。<br>

## 連線至 Exchange Online PowerShell

微軟在 2021 年 9 月的時候，公告了在 2022 年 10 月 1 日起，停用在 EXO 中的基本驗證，連帶影響的有 Outlook、EWS、RPS、POP、IMAP 和 EAS，其中也包括了 Powershell EXO V2 module 所採用的基本驗證連線方式，所以本篇會以 EXO V3 Module，來連線至 Exchange Online。<br>

- 在開始執行指令連線之前，我們先指派必要的權限，開啟Exchange Online系統管理中心，至權限分類中，選擇Oranization Management，並在角色欄位中新增Address Lists。然而 Exchange Admin Center 面臨到介面變更的狀況，所以我也會提供兩種介面的截圖給大家參考。<br>
  - 新介面<br>
    ![Github](https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/permissioin-new.png)<br>
  - 傳統介面<br>
    ![Github](https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/permissioin-old.png)<br>
- 前置作業檢查：<br>
  - 如果您使用 Windows Client: Windows 7 Service Pack 1 (SP1)、Windows 8.1、Windows 10/11<br>
  - 如果您使用 Windows Server: Windows Server 2008 R2 SP1、Windows Server 2012/R2 或更新版本<br>
  - 確認您的環境已安裝 Microsoft .NET Framework 4.7.1<br>
- 使用`系統管理員身分`開啟您本機電腦的 Windows Powershell 或是 Windows Powershell ISE，依據您熟悉的工具使用即可，本篇會使用 Windows Powershell ISE，原因是在定義通訊清單的時候，有指令碼窗格會減少蠻多重複剪貼的工作。<br>
- 開啟之後，執行 `Set-ExecutionPolicy RemoteSigned` 將 PowerShell 執行原則設定為 RemoteSigned，否則您會看到錯誤。<br>
- 完成後，執行 `Install-Module -Name ExchangeOnlineManagement` 安裝最新的 EXO 模組，也提供幾個指令可以進行選用。<br>
  - 查看目前安裝的模組版本及其安裝位置 `Get-InstalledModule ExchangeOnlineManagement | Format-List Name,Version,InstalledLocation`<br>
  - 更新模組 `Update-Module -Name ExchangeOnlineManagement`<br>
- 輸入 `Connect-ExchangeOnline -UserPrincipalName brian@M365B259147.onmicrosoft.com` 連線至 EXO，請將 UserPrincipalName 後面的帳號，輸入具有權限的帳號，通常來說可能是全域管理員，執行後會跳出帳號密碼登入的視窗，登入即可。<br>
 ![Github](/https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/login1.png)<br>
- 執行`Get-EXOMailbox`確認是否已成功連線。<br>

## 通訊清單管理

通訊清單主要是為了讓使用者可以對應企業部門的階層分類，快速找到對應的信箱使用者。在 EXO 環境中，常見的情境會有以下兩種：<br>
情境 1：信箱使用者未存在：先建立通訊清單，在建立信箱使用者。<br>
情境 2：信箱使用者已存在：建立通訊清單，更新使用者資料。例：通訊清單已部門篩選人員，使用者部門已有相對應值，必須先將其更改，再改回原本相對應值，聯絡人才會在相對應分類出現。如果環境中有使用身分混合識別，需要透過執行 AAD Sync 手動同步指令，更改過的通訊清單才會正常篩選。<br>

所以在開始執行之前，要先做好幾項前置作業：<br>
1. 確認身分識別類型是純雲端環境還是混合身分識別，純雲端可以在 M365 Admin Center 直接更改，混合身分識別需要再內部部署的網域控制上進行變更。<br>
2. 確認企業完整通訊清單階層設計。<br>
3. 確認要針對哪個屬性進行篩選，例如部門。<br>

- 新增 Address List<br>
  
  假如我們今天要建立的通訊清單階層是：<br>
  ````
    Contoso
     |-部門 1
        |-業務單位
        |-技術單位
     |-部門 2
  ````
    - 首先我們先建立第一個階層 Contoso，執行以下指令，篩選只要是使用者信箱的屬性，就會被分類近來<br>
        ````Powershell
        New-AddressList -Name Contoso -RecipientFilter {(RecipientType -eq 'UserMailbox')}
        ````
    - 再來我們要建立子階層，在 Contoso 下方新增部門 1 與部門 2，非分別對於部門 1 與部門 2的屬性進行篩選，執行以下指令<br>
        ````Powershell
        New-AddressList -Name "部門1"  -Container \Contoso -RecipientFilter {( Department -eq 'Seg1') -and (RecipientType -eq 'UserMailbox')}
        New-AddressList -Name "部門2" -Container \Contoso -RecipientFilter {( Department -eq 'Seg2') -and (RecipientType -eq 'UserMailbox')}
        ````
    - 在部門 1 在建立業務單位與技術單位巢狀階層，執行以下指令<br>
        ````Powershell
        New-AddressList -Name "業務單位" -Container \Contoso\部門1 -RecipientFilter {( Department -eq 'Seg1-sales') -and (RecipientType -eq 'UserMailbox')}
        New-AddressList -Name "技術單位" -Container \Contoso\部門1 -RecipientFilter {( Department -eq 'Seg1-tech') -and (RecipientType -eq 'UserMailbox')}
        ````
    - 完成後的樹狀結構結果呈現<br>
      - Powershell 結果視窗
        ![Github](/https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/new-addresslist.png)<br>
      - Outlook Cient User View<br>
        ![Github](https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/addresslist-outlook-view.png)<br>
    - 測試更改使用者部門，觀察篩選結果<br>
      - 假設會篩選屬於部門 1 的業務人員，將 Alex、Alan 的部門輸入 Seg1-sales。<br>
        ![Github](/https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/address-list-show.png)<br>
- 查詢 Address List<br>
  - 輸入此命令 `Get-addresslist` 會查詢通訊清單，結果如下圖。<br>
    ![Github](/https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/get-address-list.png)<br>
- 更改 Address List 名稱<br>
  - 如果要將原先的技術單位階層，更改為客服單位，請輸入以下指令，將`技術單位`改成`客服單位`，結果如下圖。<br>
    ````Powershell
    Set-AddressList -Identity "技術單位" -Name 客服單位 -DisplayName 客服單位
    ````
    ![Github](https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/set-address-list-name.png)<br>
- 更改 Address List 過濾屬性<br>
  - 剛剛只有將名稱更改，但這樣沒辦法篩選到所對應的客服單位的人員，假設客服單位的部門篩選屬性為 `Seg1-help`，請輸入以下指令，結果如下圖。<br>
    ````Powershell
    Set-AddressList -Identity "客服單位" -RecipientFilter {( Department -eq 'Seg1-help') -and (RecipientType -eq 'UserMailbox')}
    ````
    ![Github](/https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/set-address-list-filter.png)<br>
- 刪除 Address List<br>
  - 有時可能一時做太快，刪除重新建立可以是更快速的做法，假設要刪除客服單位，請輸入以下指令，結果如下圖。<br>
    ````Powershell
    Remove-AddressList -Identity "客服單位"
    ````
    ![Github](https://github.com/BrianHsing/Manage-EXO-AddressList/blob/main/images/remove-address-list.png)<br>

## 參考來源

- [Basic Authentication and Exchange Online – September 2021 Update](https://techcommunity.microsoft.com/t5/exchange-team-blog/basic-authentication-and-exchange-online-september-2021-update/ba-p/2772210)<br>
- [Manage address lists in Exchange Online](https://learn.microsoft.com/en-us/exchange/address-books/address-lists/manage-address-lists)<br>