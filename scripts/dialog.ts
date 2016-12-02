const dialog: { close: () => void } = VSS.getConfiguration();

let counter = 10;

function cancelAutoClose() {
    clearInterval(id);
    $("#message").hide();
}

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
