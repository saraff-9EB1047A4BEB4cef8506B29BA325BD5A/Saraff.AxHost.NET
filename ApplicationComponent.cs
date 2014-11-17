/* Этот файл является частью библиотеки Saraff.AxHost.NET
 * © SARAFF SOFTWARE (Кирножицкий Андрей), 2014.
 * Saraff.AxHost.NET - свободная программа: вы можете перераспространять ее и/или
 * изменять ее на условиях Меньшей Стандартной общественной лицензии GNU в том виде,
 * в каком она была опубликована Фондом свободного программного обеспечения;
 * либо версии 3 лицензии, либо (по вашему выбору) любой более поздней
 * версии.
 * Saraff.AxHost.NET распространяется в надежде, что она будет полезной,
 * но БЕЗО ВСЯКИХ ГАРАНТИЙ; даже без неявной гарантии ТОВАРНОГО ВИДА
 * или ПРИГОДНОСТИ ДЛЯ ОПРЕДЕЛЕННЫХ ЦЕЛЕЙ. Подробнее см. в Меньшей Стандартной
 * общественной лицензии GNU.
 * Вы должны были получить копию Меньшей Стандартной общественной лицензии GNU
 * вместе с этой программой. Если это не так, см.
 * <http://www.gnu.org/licenses/>.)
 * 
 * This file is part of Saraff.AxHost.NET.
 * © SARAFF SOFTWARE (Kirnazhytski Andrei), 2014.
 * Saraff.AxHost.NET is free software: you can redistribute it and/or modify
 * it under the terms of the GNU Lesser General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * (at your option) any later version.
 * Saraff.AxHost.NET is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
 * GNU Lesser General Public License for more details.
 * You should have received a copy of the GNU Lesser General Public License
 * along with Saraff.AxHost.NET. If not, see <http://www.gnu.org/licenses/>.
 * 
 * PLEASE SEND EMAIL TO:  AxHost@saraff.ru.
 */
using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Collections.ObjectModel;

namespace Saraff.AxHost {

    /// <summary>
    /// Базовый класс для пользовательских компонентов, требующих хостинг в приложениях c неуправляемым кодом.
    /// </summary>
    [ComVisible(false)]
    public class ApplicationComponent:Component {

        /// <summary>
        /// Initializes a new instance of the <see cref="ApplicationComponent"/> class.
        /// </summary>
        public ApplicationComponent() {
        }

        internal void _Initialize(ReadOnlyCollection<object> args) {
            this.Construct(args);
        }

        /// <summary>
        /// Выполняет инициализацию пользовательского элемента управления.
        /// </summary>
        /// <param name="args">Коллекция параметров.</param>
        protected virtual void Construct(ReadOnlyCollection<object> args) {
        }
    }
}
