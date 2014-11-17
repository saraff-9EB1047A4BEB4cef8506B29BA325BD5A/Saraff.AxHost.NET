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
using System.Globalization;

namespace Saraff.AxHost {

    /// <summary>
    /// Описание метода.
    /// </summary>
    [ComVisible(true)]
    [ProgId("Saraff.AxHost.MethodDescriptor")]
    [Guid("ABD68B7E-EE21-4F15-8325-E2136E6716E7")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IMethodDescriptor))]
    [Description("Описание метода.")]
    public sealed class MethodDescriptor:IMethodDescriptor {
        private List<object> _params=new List<object>();

        /// <summary>
        /// Initializes a new instance of the <see cref="MethodDescriptor"/> class.
        /// </summary>
        /// <param name="name">The name.</param>
        public MethodDescriptor(string name) {
            this.MethodName=name;
        }

        #region IMethodDescriptor Members

        /// <summary>
        /// Имя метода.
        /// </summary>
        /// <value></value>
        public string MethodName {
            get;
            private set;
        }

        /// <summary>
        /// Добавляет параметр.
        /// </summary>
        /// <param name="value">Значение параметра.</param>
        /// <param name="type">Тип параметра.</param>
        public void AddParam(object value, ParamType type) {
            switch(type) {
                case ParamType.Bool:
                    this._params.Add(Convert.ToBoolean(value, CultureInfo.InvariantCulture));
                    break;
                case ParamType.Char:
                    this._params.Add(Convert.ToChar(value, CultureInfo.InvariantCulture));
                    break;
                case ParamType.DateTime:
                    this._params.Add(Convert.ToDateTime(value, CultureInfo.InvariantCulture));
                    break;
                case ParamType.Double:
                    this._params.Add(Convert.ToDouble(value, CultureInfo.InvariantCulture));
                    break;
                case ParamType.Int:
                    this._params.Add(Convert.ToInt32(value, CultureInfo.InvariantCulture));
                    break;
                case ParamType.String:
                    this._params.Add(Convert.ToString(value, CultureInfo.InvariantCulture));
                    break;
                case ParamType.Unknown:
                default:
                    this._params.Add(value);
                    break;
            }
        }

        /// <summary>
        /// Количество параметров.
        /// </summary>
        /// <value></value>
        public int Count {
            get {
                return this._params.Count;
            }
        }

        /// <summary>
        /// Возвращает значение параметра.
        /// </summary>
        /// <value></value>
        /// <returns>Значение параметра.</returns>
        public object this[int index] {
            get {
                return this._params[index];
            }
        }

        #endregion

        internal sealed class TypeIsObject {

            internal TypeIsObject(object value) {
                this.Value=value;
            }

            internal object Value {
                get;
                private set;
            }
        }
    }

    /// <summary>
    /// Представляет собой описание метода.
    /// </summary>
    [ComVisible(true)]
    [Guid("1CC9FC96-15B6-4571-B8B8-3002A6517679")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IMethodDescriptor {

        /// <summary>
        /// Имя метода.
        /// </summary>
        [DispId(0x60020001)]
        string MethodName {
            [Description("Имя метода.")]
            get;
        }

        /// <summary>
        /// Добавляет параметр.
        /// </summary>
        /// <param name="value">Значение параметра.</param>
        /// <param name="type">Тип параметра.</param>
        [DispId(0x60020002)]
        [Description("Добавляет параметр.")]
        void AddParam(object value, ParamType type);

        /// <summary>
        /// Количество параметров.
        /// </summary>
        [DispId(0x60020003)]
        int Count {
            [Description("Количество параметров.")]
            get;
        }

        /// <summary>
        /// Возвращает значение параметра.
        /// </summary>
        /// <param name="index">Индес параметра.</param>
        /// <returns>Значение параметра.</returns>
        [DispId(0x60020004)]
        object this[int index] {
            [Description("Возвращает значение параметра.")]
            get;
        }
    }

    /// <summary>
    /// Типы параметров.
    /// </summary>
    [ComVisible(true)]
    [Guid("CB73C100-5462-46DC-AFF1-E7B37845568E")]
    public enum ParamType {

        /// <summary>
        /// Целое число.
        /// </summary>
        Int=1,

        /// <summary>
        /// Числос плавающей точкой.
        /// </summary>
        Double=2,

        /// <summary>
        /// Строка.
        /// </summary>
        String=3,

        /// <summary>
        /// Символ.
        /// </summary>
        Char=4,

        /// <summary>
        /// Булево значение.
        /// </summary>
        Bool=5,

        /// <summary>
        /// Дата/время.
        /// </summary>
        DateTime=6,

        /// <summary>
        /// Неизвесный тип.
        /// </summary>
        Unknown=7
    }
}
