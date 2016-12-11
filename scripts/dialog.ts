const dialog: { close: () => void } = VSS.getConfiguration();

let counter = 15;

const id = setInterval(() => {
    counter--;
    if (counter < 0) {
        dialog.close();
        clearInterval(id);
    }
    else {
        $("#countdown").text(counter);
    }
}, 1000);

function cancelAutoClose() {
    clearInterval(id);
    $("#message").hide();
}

function openUrl(url: string) {
    // If you clicked a link, we will cancel auto close...
    cancelAutoClose();
    VSS.getService(VSS.ServiceIds.Navigation).then((navigationService: any) => {
        navigationService.openNewWindow(url);
    });
}
