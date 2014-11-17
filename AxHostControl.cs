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
using System.IO;
using System.Diagnostics;
using Saraff.AxHost.Core;
using System.Drawing;
using System.Collections.ObjectModel;
using System.Windows.Forms;

namespace Saraff.AxHost {

    /// <summary>
    /// Предоставляет хостинг для пользовательских элементов управления в приложениях c неуправляемым кодом.
    /// </summary>
    [ComVisible(true)]
    [ProgId("Saraff.AxHost.AxHostControl")]
    [Guid("83368D28-D24A-4543-9B1A-6D7B0152C54F")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComDefaultInterface(typeof(IAxHost))]
    [ComSourceInterfaces(typeof(IAxHostEvents))]
    [DefaultProperty("ApplicationControlName")]
    [DefaultEvent("FireEvent")]
    [Description("Предоставляет хостинг для пользовательских элементов управления в приложениях c неуправляемым кодом.")]
    public sealed class AxHostControl:AxHostBaseControl, IAxHost {
        private string _applicationTypeName;
        private IContainer _components=new Container();
        private ComponentContainer<ApplicationControlAttribute> _container=null;
        private List<object> _componentParams=new List<object>();

        /// <summary>
        /// Initializes a new instance of the <see cref="AxHostControl"/> class.
        /// </summary>
        public AxHostControl() {
            this.Font=new Font("Microsoft Sans Serif", 8.25F, FontStyle.Regular, GraphicsUnit.Point, ((byte)(204)));
            this.BackColor=SystemColors.Control;

            this._container=new ComponentContainer<ApplicationControlAttribute>(this._components);
            this._container.FireEvent+=_container_FireEvent;
        }

        #region IAxHost Members

        /// <summary>
        /// Возвращает или устанавливает имя пользовательского элемента управления.
        /// </summary>
        /// <value></value>
        /// <remarks>Имя имеет формат: "имя файла сборки"!"полное имя элемента управления"</remarks>
        /// <example>MyAssembly.dll!MyNamespace.MyControl</example>
        [DefaultValue("")]
        [Category("Behavior")]
        public string ApplicationTypeName {
            get {
                return this._applicationTypeName;
            }
            set {
                try {
                    if(value==null) {
                        this._applicationTypeName=null;
                    } else {
                        string[] _args=value.Trim().Split(new char[] { '!' }, 2, StringSplitOptions.None);
                        if(_args.Length==2) {
                            this._applicationTypeName=value;
                        }
                    }
                } catch(Exception ex) {
                    Debug.WriteLine(string.Format("{1}: {2}{0}{3}", Environment.NewLine, ex.GetType().Name, ex.Message, ex.StackTrace));
                }
            }
        }

        /// <summary>
        /// Возвращает или устанавливает рабочий каталог.
        /// </summary>
        /// <value></value>
        public string WorkingDirectory {
            get {
                return this._container.WorkingDirectory;
            }
            set {
                this._container.WorkingDirectory=new Uri(value).AbsoluteUri;
            }
        }

        /// <summary>
        /// Создает и возвращает описание метода.
        /// </summary>
        /// <param name="methodName">Имя метода.</param>
        /// <returns>Описание метода.</returns>
        public MethodDescriptor CreateMethodDescriptor(string methodName) {
            try {
                return new MethodDescriptor(methodName);
            } catch(Exception ex) {
                Debug.WriteLine(string.Format("{1}: {2}{0}{3}", Environment.NewLine, ex.GetType().Name, ex.Message, ex.StackTrace));
            }
            return null;
        }

        /// <summary>
        /// Выполняет указанный метод и возвращает результат.
        /// </summary>
        /// <param name="methodDescriptor">Описание метода.</param>
        /// <returns>Результат.</returns>
        public object PerformMethod(MethodDescriptor methodDescriptor) {
            try {
                return this._container.PerformMethod(methodDescriptor);
            } catch(Exception ex) {
                Debug.WriteLine(string.Format("{1}: {2}{0}{3}", Environment.NewLine, ex.GetType().Name, ex.Message, ex.StackTrace));
            }
            return null;
        }

        /// <summary>
        /// Добавляет параметр компонента.
        /// </summary>
        /// <param name="param">Значение параметра.</param>
        public void AddComponentParameter(object param) {
            try {
                this._componentParams.Add(param);
            } catch(Exception ex) {
                Debug.WriteLine(string.Format("{1}: {2}{0}{3}", Environment.NewLine, ex.GetType().Name, ex.Message, ex.StackTrace));
            }
        }

        /// <summary>
        /// Загружает и инициализирует компонент.
        /// </summary>
        public new void Load() {
            this._container.CreateComponentInstance(this.ApplicationTypeName);
            if(this._container.ApplicationComponent!=null) {
                this.Controls.Add((Control)this._container.ApplicationComponent);
                ((Control)this._container.ApplicationComponent).Dock=DockStyle.Fill;
                ((ApplicationControl)this._container.ApplicationComponent)._Initialize(this._componentParams.AsReadOnly());
            }
        }

        #endregion

        #region IAxHostEvents Members

        /// <summary>
        /// Возникает в момент, когда необходимо обработать одно из событий пользовательского элемента управления.
        /// </summary>
        [Category("Action")]
        public event AxHostControlFireEventHandler FireEvent;

        #endregion

        private void _container_FireEvent(object sender, ComponentContainer<ApplicationControlAttribute>.FireEventArgs e) {
            try {
                if(FireEvent!=null) {
                    this.FireEvent(e.EventId);
                }
            } catch(Exception ex) {
                Debug.WriteLine(string.Format("{1}: {2}{0}{3}", Environment.NewLine, ex.GetType().Name, ex.Message, ex.StackTrace));
            }
        }
    }

    /// <summary>
    /// Представляет метод обработчика события FireEvent.
    /// </summary>
    /// <param name="eventId">Описание события.</param>
    [ComVisible(false)]
    public delegate void AxHostControlFireEventHandler(EventDescriptor eventId);
}
