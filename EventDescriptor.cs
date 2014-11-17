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
using System.Reflection;
using System.Diagnostics;

namespace Saraff.AxHost {

    /// <summary>
    /// Описание события.
    /// </summary>
    [ComVisible(true)]
    [ProgId("Saraff.AxHost.EventDescriptor")]
    [Guid("548F0EF9-DA93-4CBB-8DC3-91EF5F35708B")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IEventDescriptor))]
    [Description("Описание события.")]
    public sealed class EventDescriptor:IEventDescriptor {
        private string _event_name;
        private EventArgs _event_args;
        private Dictionary<string, object> _params=new Dictionary<string, object>();

        /// <summary>
        /// Initializes a new instance of the <see cref="EventDescriptor"/> class.
        /// </summary>
        /// <param name="eventName">Name of the event.</param>
        /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
        public EventDescriptor(string eventName, EventArgs e) {
            this._event_name=eventName;
            this._event_args=e;
            foreach(PropertyInfo _property in e.GetType().GetProperties(BindingFlags.Public|BindingFlags.Instance)) {
                this._params.Add(_property.Name, _property.GetValue(e, null));
            }
        }

        #region IEventDescriptor Members

        /// <summary>
        /// Возвращает имя события.
        /// </summary>
        /// <value></value>
        public string EventName {
            get {
                return this._event_name;
            }
        }

        /// <summary>
        /// Возвращает количество параметров.
        /// </summary>
        /// <value></value>
        public int ParamCount {
            get {
                return this._params.Count;
            }
        }

        /// <summary>
        /// Возвращает значение указанного параметра.
        /// </summary>
        /// <param name="paramName">Имя параметра.</param>
        /// <returns>Значение параметра.</returns>
        public object GetParam(string paramName) {
            try {
                return this._params[paramName];
            } catch(Exception ex) {
                Debug.WriteLine(string.Format("{1}: {2}{0}{3}", Environment.NewLine, ex.GetType().Name, ex.Message, ex.StackTrace));
            }
            return null;
        }

        /// <summary>
        /// Устанавливает значение указанного параметра.
        /// </summary>
        /// <param name="paramName">Имя параметра.</param>
        /// <param name="value">Значение параметра.</param>
        /// <returns>
        /// true, если значение установлено успешно; иначе, false.
        /// </returns>
        public bool SetParam(string paramName, object value) {
            try {
                PropertyInfo _property=this._event_args.GetType().GetProperty(paramName);
                if(_property!=null&&_property.GetSetMethod()!=null) {
                    object _value=Convert.ChangeType(value, _property.PropertyType);
                    _property.SetValue(this._event_args, _value, null);
                    this._params[paramName]=_value;
                    return true;
                }
            } catch(Exception ex) {
                Debug.WriteLine(string.Format("{1}: {2}{0}{3}", Environment.NewLine, ex.GetType().Name, ex.Message, ex.StackTrace));
            }
            return false;
        }

        #endregion

        /// <summary>
        /// Возвращает массив параметров компонента.
        /// </summary>
        public string[] Params {
            get {
                List<string> _params_names=new List<string>();
                foreach(string _param in this._params.Keys) {
                    _params_names.Add(_param);
                }
                return _params_names.ToArray();
            }
        }
    }

    /// <summary>
    /// Представляет собой описание события.
    /// </summary>
    [ComVisible(true)]
    [Guid("50B33711-F2AB-4537-856F-454BF3237EC8")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IEventDescriptor {

        /// <summary>
        /// Возвращает имя события.
        /// </summary>
        [DispId(0x60020001)]
        string EventName {
            [Description("Возвращает имя события.")]
            get;
        }

        /// <summary>
        /// Возвращает количество параметров.
        /// </summary>
        [DispId(0x60020002)]
        int ParamCount {
            [Description("Возвращает количество параметров.")]
            get;
        }

        /// <summary>
        /// Возвращает значение указанного параметра.
        /// </summary>
        /// <param name="paramName">Имя параметра.</param>
        /// <returns>Значение параметра.</returns>
        [DispId(0x60020003)]
        [Description("Возвращает значение указанного параметра.")]
        object GetParam(string paramName);

        /// <summary>
        /// Устанавливает значение указанного параметра.
        /// </summary>
        /// <param name="paramName">Имя параметра.</param>
        /// <param name="value">Значение параметра.</param>
        /// <returns>true, если значение установлено успешно; иначе, false.</returns>
        [DispId(0x60020004)]
        [Description("Устанавливает значение указанного параметра.")]
        bool SetParam(string paramName, object value);
    }
}
