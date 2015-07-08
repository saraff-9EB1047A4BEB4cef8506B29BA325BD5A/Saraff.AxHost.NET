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
using System.Runtime.InteropServices;
using System.ComponentModel;

namespace Saraff.AxHost {

    /// <summary>
    /// Интерфейс компонента хостинга для пользовательских элементов управления в приложениях c неуправляемым кодом.
    /// </summary>
    [ComVisible(true)]
    [Guid("3A0972A6-BF82-440A-8097-D0B8F802B5C8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IAxHost:IDisposable {

        /// <summary>
        /// Возвращает или устанавливает имя пользовательского элемента управления.
        /// </summary>
        /// <remarks>Имя имеет формат: "имя файла сборки"!"полное имя элемента управления"</remarks>
        /// <example>MyAssembly.dll!MyNamespace.MyControl</example>
        [DispId(0x60020001)]
        string ApplicationTypeName {
            [Description("Возвращает или устанавливает имя пользовательского элемента управления.")]
            get;
            set;
        }

        /// <summary>
        /// Возвращает или устанавливает рабочий каталог.
        /// </summary>
        [DispId(0x60020002)]
        string WorkingDirectory {
            [Description("Возвращает или устанавливает рабочий каталог.")]
            get;
            set;
        }

        /// <summary>
        /// Создает и возвращает описание метода.
        /// </summary>
        /// <param name="methodName">Имя метода.</param>
        /// <returns>Описание метода.</returns>
        [DispId(0x60020003)]
        [Description("Создает и возвращает описание метода.")]
        MethodDescriptor CreateMethodDescriptor(string methodName);

        /// <summary>
        /// Выполняет указанный метод и возвращает результат.
        /// </summary>
        /// <param name="methodDescriptor">Описание метода.</param>
        /// <returns>Результат.</returns>
        [DispId(0x60020004)]
        [Description("Выполняет указанный метод и возвращает результат.")]
        object PerformMethod(MethodDescriptor methodDescriptor);

        /// <summary>
        /// Добавляет параметр компонента.
        /// </summary>
        /// <param name="param">Значение параметра.</param>
        [DispId(0x60020005)]
        [Description("Добавляет параметр компонента.")]
        void AddComponentParameter(object param);

        /// <summary>
        /// Загружает и инициализирует компонент.
        /// </summary>
        [DispId(0x60020006)]
        [Description("Загружает и инициализирует компонент, указанный в ApplicationTypeName.")]
        void Load();

        /// <summary>
        /// Освобождает ресурсы занятые компонентом.
        /// </summary>
        [DispId(0x60020007)]
        [Description("Освобождает ресурсы занятые компонентом.")]
        void Dispose();
    }
}
