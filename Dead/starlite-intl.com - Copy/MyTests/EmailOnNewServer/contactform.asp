<!DOCTYPE html>

<!-- 
    [BN, 8/11/15] Test of using new email system on NEW server at HostMySite.com, based on Knowledgebase article at
    https://solutions.hostmysite.com/index.php?/Knowledgebase/Article/View/8596/0/Using-CDOSys-to-create-an-ASP-Mail-form-that-uses-Authentication
    
    See also http://www.w3schools.com/asp/asp_send_email.asp where it says:
    "CDOSYS is a built-in component in ASP. This component is used to send e-mails with ASP." It replaces CDONTs.
-->


<html xmlns="http://www.w3.org/1999/xhtml">


<body>
    <form name="contactform" method="post" action="send_form_email.asp">
        <table width="450px">
            <tr>
                <td valign="top">
                    <label for="first_name">First Name *</label>
                </td>
                <td valign="top">
                    <input type="text" name="first_name" maxlength="50" size="30">
                </td>
            </tr>
            <tr>
                <td valign="top" ">
                    <label for=" last_name">Last Name *</label>
                </td>
                <td valign="top">
                    <input type="text" name="last_name" maxlength="50" size="30">
                </td>
            </tr>
            <tr>
                <td valign="top" ">
                    <label for=" home_address">Home Address</label>
                </td>
                <td valign="top">
                    <textarea name="home_address" maxlength="100" cols="25 Rows=" 6"></textarea>
                </td>
            </tr>
            <tr>
                <td valign="top">
                    <label for="email">Email Address *</label>
                </td>
                <td valign="top">
                    <input type="text" name="email" maxlength="80" size="30">
                </td>
            </tr>
            <tr>
                <td valign="top">
                    <label for="telephone">Telephone Number</label>
                </td>
                <td valign="top">
                    <input type="text" name="telephone" maxlength="30" size="30">
                </td>
            </tr>
            <tr>
                <td valign="top">
                    <label for="comments">Comments *</label>
                </td>
                <td valign="top">
                    <textarea name="comments" maxlength="1000" cols="25" rows="6"></textarea>
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align:center">
                    * - Denotes a required field
                    <input type="submit" value="Submit">
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
