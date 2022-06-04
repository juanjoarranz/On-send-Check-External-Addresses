/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/* eslint-disable office-addins/no-office-initialize */

// The Office initialize function must be run each time a new page is loaded.
Office.initialize = function (reason) {
    document.querySelector('#send-approved').addEventListener('click', () => sendMessage("Send"));
    document.querySelector('#send-rejected').addEventListener('click', () => sendMessage("Do not Send"));

    let section = document.querySelector('section');
    window.addEventListener('resize', () => alignVertically(section));
    alignVertically(section);
};

function sendMessage(message) {
    Office.context.ui.messageParent(message);
}

function alignVertically(section: HTMLElement) {
    section.style.marginTop = Math.round((window.innerHeight - section.offsetHeight) / 2) - 15 + 'px';
    // console.log("window.innerHeight = " + window.innerHeight);
}