
Office.onReady(async (info) => {
    console.log("Office.onReady called");

    let emailSaveStatusBanner = document.getElementById("emailSaveStatusBanner")

    if (info.host === Office.HostType.Outlook) {
        await checkEmailSaveStatus(emailSaveStatusBanner);
    }
});


async function checkEmailSaveStatus(emailSaveStatusBanner) {
    console.log("Checking email save status");

    emailSaveStatusBanner.textContent  = "Checking email save status...";

    if (true) {
        emailSaveStatusBanner.textContent  = "Email has been saved";
        return;
    }

    emailSaveStatusBanner.textContent  = "Email has not been saved";
    return;

}
