if (typeof IntersectionObserver !== 'undefined') {
    document.documentElement.classList.add('js');
}
document.addEventListener('DOMContentLoaded', function () {
    var els = document.querySelectorAll('.reveal');
    if (typeof IntersectionObserver === 'undefined') {
        els.forEach(function (el) { el.classList.add('in'); });
        return;
    }
    var io = new IntersectionObserver(function (entries) {
        entries.forEach(function (e) {
            if (e.isIntersecting) {
                e.target.classList.add('in');
                io.unobserve(e.target);
            }
        });
    }, { threshold: 0.12 });
    els.forEach(function (el) { io.observe(el); });
});

function toggleMenu() {
    var m = document.getElementById('mobile-menu');
    m.style.display = m.style.display === 'flex' ? 'none' : 'flex';
}
