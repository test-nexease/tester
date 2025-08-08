import streamlit as st
import pandas as pd
import win32com.client as win32

st.title("Bulk Email Sender with Outlook and Excel Table Snippet")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.write("Preview of uploaded data:")
    st.dataframe(df)

    required_cols = {'E Mail ID', 'Department', 'Comment', 'Status', 'Supplier No', 'Supplier Name',
                     'Purchase Order No', 'Item', 'Purchase Order Date', 'Material', 'Short Text',
                     'Order Quantity', 'Order Unit', 'Unit Price', 'Order Amount', 'Pending Qty',
                     'Pending Amount', 'Storage location', 'PR No', 'End User'}

    # Add 'CC' column as optional
    if 'CC' not in df.columns:
        st.warning("No 'CC' column found. Proceeding without CC recipients.")

    if not required_cols.issubset(set(df.columns)):
        missing = required_cols - set(df.columns)
        st.error(f"Missing columns in Excel file: {missing}")
    else:
        if st.button("Send Emails"):
            outlook = win32.Dispatch('outlook.application')
            success_count = 0
            fail_count = 0

            grouped = df.groupby('Supplier Name')

            for email, group in grouped:
                try:
                    # Combine CC emails from all rows for this email group (unique and joined by semicolon)
                    if 'CC' in df.columns:
                        cc_emails = group['CC'].dropna().unique()
                        cc_emails = [cc for cc in cc_emails if cc.strip() != '']  # filter out empty strings
                        cc_string = "; ".join(cc_emails) if cc_emails else ""
                    else:
                        cc_string = ""

                    rows_html = ""
                    for _, row in group.iterrows():
                        rows_html += f"""
                        <tr>
                            <td>{row['Department']}</td>
                            <td>{row['Comment']}</td>
                            <td>{row['Status']}</td>
                            <td>{row['Supplier No']}</td>
                            <td>{row['Supplier Name']}</td>
                            <td>{row['Purchase Order No']}</td>
                            <td>{row['Item']}</td>
                            <td>{row['Purchase Order Date'].strftime('%Y-%m-%d') if not pd.isna(row['Purchase Order Date']) else ''}</td>
                            <td>{row['Material']}</td>
                            <td>{row['Short Text']}</td>
                            <td>{row['Order Quantity']}</td>
                            <td>{row['Order Unit']}</td>
                            <td>{row['Unit Price']}</td>
                            <td>{row['Order Amount']}</td>
                            <td>{row['Pending Qty']}</td>
                            <td>{row['Pending Amount']}</td>
                            <td>{row['Storage location']}</td>
                            <td>{row['PR No']}</td>
                            <td>{row['End User']}</td>
                        </tr>
                        """

                    html_table = f"""
                    <table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; font-family: Arial, sans-serif;">
                        <tr style="background-color: #f2f2f2;">
                            <th>Department</th><th>Comment</th><th>Status</th><th>Supplier No</th><th>Supplier Name</th>
                            <th>Purchase Order No</th><th>Item</th><th>Purchase Order Date</th><th>Material</th><th>Short Text</th>
                            <th>Order Quantity</th><th>Order Unit</th><th>Unit Price</th><th>Order Amount</th><th>Pending Qty</th>
                            <th>Pending Amount</th><th>Storage location</th><th>PR No</th><th>End User</th>
                        </tr>
                        {rows_html}
                    </table>
                    """

                    mail = outlook.CreateItem(0)
                    mail.To = email
                    if cc_string:
                        mail.CC = cc_string
                    mail.Subject = "Rock Tools Pune Delivery Follow-up – Action Required Before Material Dispatch"
                    mail.HTMLBody = f"""
                    <p>Hello,<br>Greetings!!</p>
                    <p>Please go through the below pending orders as per our system. <b><span style="background-color: yellow;">"The yellow-marked entries"</span></b> indicate the pending quantity.<br>
                    You are requested to discuss the same with the respective end user and <b>arrange to dispatch the material only after receiving confirmation from the end user.</b><br>
                    Kindly note that <b>any material sent without proper communication</b> (Verbal / Message / WhatsApp / Email confirmation from the end user) <b>will not be accepted.</b><br>
                    This email is intended as part of our delivery follow-up and to share the current open PO status for your reference.
                    </p>
                    {html_table}
                    <br><br>
                    <p>Best Regards,<br>
                    <b>Pratheeshkumar Nair</b><br>
                    Deputy Manager-Indirect Purchase,<br>
                    <b>Sandvik Mining and Rock Technology India Private Limited</b><br>
                    Mumbai Pune Road, Dapodi, Pune - 411 012, India<br>
                    Tel No.: +91 20 63102209, Mob (O): +91 9923699610<br>
                    Email: pratheesh.kumar@sandvik.com</p>
                    """

                    mail.Send()
                    success_count += 1
                except Exception as e:
                    st.error(f"Failed to send email to {email}: {e}")
                    fail_count += 1

            st.success(f"Emails sent successfully: {success_count}")
            if fail_count > 0:
                st.warning(f"Emails failed: {fail_count}")
