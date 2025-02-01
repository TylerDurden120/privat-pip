export function showFlowWidget() {
    hideSpinner();
    document.getElementById('flow-div')!.style.display = 'block';
}

export function hideFlowWidget() {
    document.getElementById('flow-div')!.style.display = 'none';
}

function hideSpinner() {
    document.getElementById('spinner')!.style.display = 'none';
    document.getElementById('spinner-container')!.style.display = 'none';
}
