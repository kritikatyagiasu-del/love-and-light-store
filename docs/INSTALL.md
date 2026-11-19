# Install (about 10 minutes)

Runs free on a Google account: Google Sheet + Apps Script + Google Drive + Gmail.
You can install it on your own account to test, then copy it to the shop owner's account.

Files: `dist/Code.gs` and `dist/appsscript.json`.

## 1. The Sheet
1. In Google Drive, make a folder for the shop, open it, and choose **New ▸ Google Sheets ▸ Blank spreadsheet**. Name it, for example, **Store**.
2. Setup creates two folders next to the Sheet: **Store – Product photos** and **Store – Payment screenshots (private)**. If a folder with one of these names is already there, setup reuses it. Keep the screenshots folder private (don't share it).

## 2. Paste the code
1. In the Sheet: **Extensions ▸ Apps Script**.
2. Open **Code.gs**, select everything, and replace it with the contents of `dist/Code.gs`.
3. **Project Settings (gear)** ▸ tick **Show "appsscript.json" manifest file in editor** ▸ back to **Editor**.
4. Replace **appsscript.json** with the contents of `dist/appsscript.json`.
5. **Ctrl/Cmd + S**.

## 3. Run setup once
1. Choose **setup** in the function dropdown ▸ **Run** ▸ **Review permissions** ▸ your account ▸ **Advanced ▸ Go to (project) ▸ Allow**.
   It asks for Sheets, Drive (photos and screenshots), Gmail (order emails) and a timer.
2. Setup creates the **Settings / Dashboard / Products / Orders / Reviews / Log** tabs, finds or creates the two Drive folders, makes an admin PIN, and starts a 5-minute timer.
   It is safe to run again: it keeps settings, PIN, folders and data.

## 4. Publish
1. **Deploy ▸ New deployment ▸ gear ▸ Web app**.
2. **Execute as: Me**, **Who has access: Anyone** ▸ **Deploy**.
3. Copy the **Web app URL** (ends in `/exec`).

## 5. Fill in Settings (Sheet ▸ Settings tab)
| Setting | Default | Notes |
|---|---|---|
| appUrl | — | The `/exec` URL. Needed for the links in emails |
| adminEmail | your email | Where order alerts go |
| adminPhone | — | Shown to customers for help |
| pickupAddress | — | "Collect in person" is hidden until this is filled |
| upiId / upiName | — | UPI orders are refused until upiId is filled |
| shippingFee / freeShippingFrom | 10 / 500 | Delivery fee in rupees, charged below the free-delivery threshold |
| shipAreaName | Delhi NCR | Shown on "Deliver to me" |
| shipCities | Delhi, Noida, Gurugram (Gurgaon), Faridabad | The city dropdown |
| shipPinPrefixes | 110, 201, 122, 121 | A delivery PIN must start with one of these |
| holdMinutes | 30 | Items held while a customer pays by UPI |
| pickupHoldHours | 72 | Cash pickup orders kept aside this long, then released |
| proofRetentionDays | 90 | Payment screenshots go to the Drive bin after this (Google empties the bin 30 days later) |
| checkNudgeHours | 3 | Second email to the admin if a payment is still unchecked |
| categories | Candles, Gift boxes, Other crafts | Comma separated |
| lowStockAt | 2 | "Only N left" appears at or below this |
| adminAlerts | all | all / changes (payments and reviews to check, and problems only) / none |
| reviewsOn | yes | no = no review invites, no reviews in the shop |
| reviewMinForStars | 3 | A product's average stars show only from this many published reviews |
| reviewConsentText | see Sheet | The tick box on the review form |
| reviewRules | see Sheet | Shown on the review form: what gets published and what doesn't |
| consentText | see Sheet | The checkout tick box. `{store}` becomes the store name |
| privacyText | DRAFT | Replace with reviewed text. `{days}` and `{phone}` are filled in |

Reload the Sheet ▸ **Store app ▸ Show admin PIN and links** (also shows emails left today).

## 6. Use it
- **Customers:** the `/exec` link.
- **Admin:** `/exec?page=admin` on the phone ▸ PIN ▸ browser menu ▸ **Add to Home screen**. The PIN is asked again every 30 days.
- **Add products:** Products ▸ **Add product** ▸ take or choose up to 4 photos, name, category, price, how many. Use + / − on the list to change stock.
- **In-person sale:** Orders ▸ **In-person sale**. Stock goes down; no email is sent.

## Insights and dashboard
- **In the app:** the **Insights** tab has a month picker and shows money received (UPI / cash), orders, customers (total, new, returning), items sold, delivered / collected / in-person counts, payments to check, delivery fees, cancellations, cash still due, refunds due, best sellers, and what is running low right now.
- **In the Sheet:** the **Dashboard** tab shows the same numbers. Pick a month in cell **B3**; a month-by-month table sits below. It updates by itself whenever an order or stock changes, or use **Store app ▸ Refresh dashboard**. A month you pick stays until the calendar month changes, then it returns to the current month.
- Customers are counted by phone number, from paid orders. "New" means their first paid order was in that month. In-person sales with no phone number are not counted as customers.
- Don't type in the Dashboard tab other than B3: it is rewritten each time.

## Reviews
- **Who can review:** only buyers. When an order is marked **delivered** or **collected**, the customer is emailed a "Write a review" link (once per order). The order page also shows "Review what you bought". In-person sales are not invited.
- **One review per item per order**, 1 to 5 stars, text optional, with a consent tick. The buyer can choose to appear as "A verified buyer" instead of first name and initial. Reviews can't be edited after sending.
- **Nothing appears until it is checked.** The admin gets an email and a **Reviews** tab: **Publish**, **Don't publish** (with a reason), **Reply**, and **Take down** (for example when the buyer asks).
- **Rules to check by:** publish honest reviews, good or bad; reply instead of hiding. Don't publish abuse, personal details, spam, or any claim that a product treats or cures a health condition.
- **In the shop:** published reviews show under the product with "Verified purchase", the first name and initial, and any reply. The average stars show on the product and its card only once there are 3 published reviews.
- **Records:** every review stays in the Reviews tab of the Sheet with its status, reason and consent time. Customers' surnames, emails and order numbers are never shown in the shop.
- **Insights and Dashboard** count reviews received, published, not published, waiting, and the average stars.
- **Upgrading:** the privacy notice text in Settings is not changed by setup. Add a line about reviews to your own `privacyText`, for example: "If you write a review, it is shown with your first name and initial (or as a verified buyer) after we check it, and we remove it if you ask."
- Each completed online order now sends one more email (the review invite).

## Already installed an earlier version?
Paste the new `Code.gs` ▸ save ▸ run **setup** once (it adds any missing tabs (Dashboard, Reviews), the new Orders column and the new settings, and keeps everything else) ▸ **Deploy ▸ Manage deployments ▸ pencil ▸ Version: New version ▸ Deploy**.

## How an order works
1. Customer adds items ▸ cart ▸ chooses **Collect in person** (UPI or cash) or **Deliver to me** (delivery area only, UPI only) ▸ details ▸ ticks the consent box.
2. **UPI:** items are held for 30 minutes. Screen 1 shows the UPI button, QR and UPI ID. Screen 2 asks for **either** the 12-digit UPI reference number **or** a screenshot, then **I have paid**. If they haven't confirmed within 3 minutes they are emailed the link.
3. The customer sees "Your order is reserved, you'll receive an email when the order is confirmed". The items stay reserved.
4. The admin is emailed (reference number and/or screenshot) and taps **Payment received** (customer emailed, order confirmed) or **Not received** (order cancelled, items back on sale, customer emailed).
5. **Delivery:** **Mark as sent** ▸ optional delivery service, tracking link, driver name, driver phone, note ▸ customer is emailed. Details can be edited later, with or without another email. Then **Mark delivered**.
6. **Cash pickup:** confirmed at once, kept aside for 72 hours, then released automatically and both sides are emailed. On handover: **Collected · cash** or **Collected · UPI** (shows a QR).
7. **Cancel** returns the items to stock. A cancelled order that was paid moves to **Refunds to make** until **I have refunded** is tapped.
8. A payment confirmed after the 30-minute hold ended is accepted if the items are still in stock; otherwise it goes to **Refunds to make**.

## What is kept and deleted
- The Sheet is never deleted from.
- Payment screenshots live only in the private screenshots folder. Once a day the timer bins files in **that folder only** that are older than 90 days. The order row keeps its reference number.
- Product photos are shared as "anyone with the link can view" so the shop can show them.

## Testing tips
- Add one product, then order it from your phone with a second email address.
- Pay Rs.1 once (set a product to Rs.1) to check that the UPI button opens GPay/PhonePe/Paytm with the amount filled in.
- Check that product photos show in the shop. The photo link format is in one place in the code (`photo()` in `src/common.js`).
- Emails go out from the Gmail account that installed it, so only test with people who know. Free Gmail sends to about 100 recipients a day from scripts; a delivered order uses about 5 emails, including the review invite.

## Updating the code later
Paste the new `Code.gs` ▸ save ▸ run **setup** again if the notes say so ▸ **Deploy ▸ Manage deployments ▸ pencil ▸ Version: New version ▸ Deploy**. The link stays the same.

## Moving it to the shop owner's account
Share the Drive folder ▸ in their account **File ▸ Make a copy** of the Sheet into their own folder ▸ clear test rows in Products, Orders, Reviews, Log ▸ clear **adminPin**, **productFolderId** and **proofFolderId** ▸ repeat steps 3 to 6. Product photos must be added again from their admin page.

## Known limits
- No payment gateway: the admin confirms each UPI payment by hand. A screenshot or reference number can be faked; the proof is the money in the UPI app.
- Free Gmail sends to 100 recipients a day from scripts.
- Google shows a "created by a Google Apps Script user" banner on free accounts.
- Product photos are served from Google Drive. This works today but is not an officially supported way to host images.
