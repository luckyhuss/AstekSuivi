using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AstekSuivi.Model
{
    class CustomContextMenu
    {
        public CustomContextMenu(string menuType, string menuLink)
        {
            MenuType = menuType;
            MenuLink = menuLink;
        }

        public string MenuType { set; get; }
        public string MenuLink { set; get; }
    }
}
