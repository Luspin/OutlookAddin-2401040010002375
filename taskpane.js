let emailSaveStatusBanner = document.getElementById("emailSaveStatusBanner")

Office.onReady(async (info) => {
    console.log("Office.onReady called");

    if (info.host === Office.HostType.Outlook) {
        await checkEmailSaveStatus(Office.context.mailbox.item.itemId);
    }
});


async function checkEmailSaveStatus() {
    console.log("Checking email save status");

    emailSaveStatusBanner.textContent  = "Checking email save status...";

    if (true) {
        emailSaveStatusBanner.textContent  = "Email has been saved";
        return;
    }

    emailSaveStatusBanner.textContent  = "Email has not been saved";
    return;

}
