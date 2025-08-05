let backend;

window.addEventListener("DOMContentLoaded", () => {
    new QWebChannel(qt.webChannelTransport, function (channel) {
        backend = channel.objects.backend;

        document.getElementById("check-in").addEventListener("click", () => {
            backend.checkIn(); // 파이썬 슬롯 실행
        });

        document.getElementById("check-out").addEventListener("click", () => {
            backend.checkOut(); // 파이썬 슬롯 실행
        });
    });
});
