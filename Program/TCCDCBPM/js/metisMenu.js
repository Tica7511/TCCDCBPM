document.addEventListener('DOMContentLoaded', function() {
    const menuItems = document.querySelectorAll('.sidebarmenu ul li');

    menuItems.forEach(function(menuItem) {
        const submenu = menuItem.querySelector('.submenu');
        const icon = menuItem.querySelector('.icon');

        if (submenu) {
            // 如果有子选单，显示图标并添加展开/收合功能
            if (icon) {
                icon.style.display = 'inline-block';
            }

            menuItem.querySelector('.menuitemwrapper').addEventListener('click', function(e) {
                e.preventDefault(); // 阻止默认行为（如导航）
                menuItem.classList.toggle('open');
            });
        } else {
            // 如果没有子选单，隐藏图标
            if (icon) {
                icon.style.display = 'none';
            }

            // 继续导航到<a>链接
            const link = menuItem.querySelector('a');
            if (link) {
                menuItem.querySelector('.menuitemwrapper').addEventListener('click', function() {
                    if (link.target === '_blank') {
                        window.open(link.href, '_blank');
                    } else {
                        window.location.href = link.href;
                    }
                });
            }
        }
    });
});