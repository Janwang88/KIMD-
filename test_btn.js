document.addEventListener('DOMContentLoaded', () => {
    setInterval(() => {
        const btn = document.querySelector('#heroSearchBtn');
        if (btn) {
            console.log("Check button listeners:", btn.onclick);
        }
    }, 2000);
});
