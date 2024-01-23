let emailSaveStatusBanner = document.getElementById("emailSaveStatusBanner")

Office.onReady(async (info) => {
    console.log("Office.onReady called");

    if (info.host === Office.HostType.Outlook) {
        await checkEmailSaveStatus(Office.context.mailbox.item.itemId);
    }
});


async function checkEmailSaveStatus() {
    console.log("Checking email save status");

    emailSaveStatusBanner.innerText = "Checking email save status...";

    if (true) {
        emailSaveStatusBanner.innerText = "Email has been saved";
        return;
    }

    emailSaveStatusBanner.innerText = "Email has not been saved";
    return;

}
