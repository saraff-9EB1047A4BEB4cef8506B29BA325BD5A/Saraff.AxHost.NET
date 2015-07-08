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
using System.Reflection;
using System.Diagnostics;
using System.Net;

namespace Saraff.AxHost.Core {

    /// <summary>
    /// Реализует хостинг для компонентов.
    /// </summary>
    /// <typeparam name="TAttribute">Тип аттрибута, которым должны быть отмечены компоненты, размещаемые в контейнере.</typeparam>
    [ComVisible(false)]
    internal sealed class ComponentContainer<TAttribute>:Component where TAttribute:Attribute {
        private IContainer _components=new Container();
        private Component _applicationComponent=null;
        private string _appDirectory=null;

        /// <summary>
        /// Initializes a new instance of the <see cref="ComponentContainer&lt;TAttribute&gt;"/> class.
        /// </summary>
        public ComponentContainer() {
            this.WorkingDirectory=Directory.GetCurrentDirectory();
            AppDomain.CurrentDomain.ReflectionOnlyAssemblyResolve+=this._ReflectionOnlyAssemblyResolve;
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="ComponentContainer&lt;TAttribute&gt;"/> class.
        /// </summary>
        /// <param name="container">The container.</param>
        public ComponentContainer(IContainer container):this() {
            container.Add(this);
        }

        #region Открытые методы

        /// <summary>
        /// Создает экземпляр указанного компонента и устанавливает подключение к СУБД.
        /// </summary>
        /// <param name="applicationComponentName">Имя пользовательского компонента (в формате MyAssembly.dll!MyNamespace.MyComponent).</param>
        public void CreateComponentInstance(string applicationComponentName) {
            try {
                if(applicationComponentName==null) {
                    throw new ArgumentNullException("applicationComponentName", "Не указано имя компонента.");
                }
                string[] _args=applicationComponentName.Trim().Split(new char[] { '!' }, 2, StringSplitOptions.None);
                if(_args.Length!=2) {
                    throw new ArgumentException("Имя элемента управления имеет неверный формат.", "applicationComponentName");
                }

                Assembly _asm;
                if(!this._IsLoad(_args[0],out _asm)) {
                    //копируем сборку, содержащую указанный компонент
                    string _targetFile=this._CopyFile(_args[0]);

                    #region копируем файлы, указанные в аттрибуте RequiredFileAttribute

                    //т.к. необходимые для загрузки сборки зависимости еще отсутствуют,
                    //сборка загружается в контекст, предназначенный только для отражения.
                    Assembly.ReflectionOnlyLoadFrom(this.GetType().Assembly.Location);
                    _asm=Assembly.ReflectionOnlyLoadFrom(_targetFile);
                    //копируем необходимые файлы
                    foreach(CustomAttributeData _attr_data in CustomAttributeData.GetCustomAttributes(_asm)) {
                        if(_attr_data.ToString().Contains(typeof(RequiredFileAttribute).FullName)) {
                            var _reqFile=this._CopyFile((string)_attr_data.ConstructorArguments[0].Value);
                            try {
                                Assembly.ReflectionOnlyLoadFrom(_reqFile);
                            } catch {
                            }
                        }
                    }
                    for(Type _type=_asm.GetType(_args[1]); _type!=null; ) {
                        foreach(CustomAttributeData _attr_data in CustomAttributeData.GetCustomAttributes(_type)) {
                            if(_attr_data.ToString().Contains(typeof(RequiredFileAttribute).FullName)) {
                                this._CopyFile((string)_attr_data.ConstructorArguments[0].Value);
                            }
                        }
                        break;
                    }

                    #endregion

                    //загружаем сборку и создаем экземпляр компонента
                    this.ApplicationComponent=Activator.CreateInstance(Assembly.LoadFrom(_targetFile).GetType(_args[1])) as Component;
                } else {
                    this.ApplicationComponent=Activator.CreateInstance(_asm.GetType(_args[1])) as Component;
                }
            } catch(TargetInvocationException ex) {
                throw new InvalidOperationException((ex.InnerException!=null?ex.InnerException:ex).Message,ex);
            }
        }

        /// <summary>
        /// Выполняет указанный метод и возвращает результат.
        /// </summary>
        /// <param name="methodDescriptor">Описание метода.</param>
        /// <returns>
        /// Результат.
        /// </returns>
        public object PerformMethod(MethodDescriptor methodDescriptor) {
            Type[] _argsTypes=new Type[methodDescriptor.Count];
            for(int i=0; i<methodDescriptor.Count; i++) {
                _argsTypes[i]=methodDescriptor[i].GetType()==typeof(MethodDescriptor.TypeIsObject)?typeof(object):methodDescriptor[i].GetType();
            }
            MethodInfo _method=this.ApplicationComponent.GetType().GetMethod(methodDescriptor.MethodName, _argsTypes);
            if(_method!=null&&_method.GetCustomAttributes(typeof(ApplicationProcessedAttribute), false).Length>0) {
                object[] _args=new object[methodDescriptor.Count];
                for(int i=0; i<methodDescriptor.Count; i++) {
                    _args[i]=methodDescriptor[i];
                }
                return _method.Invoke(this.ApplicationComponent, _args);
            }
            return null;
        }

        #endregion

        #region Открытые свойства

        /// <summary>
        /// Возвращает экземпляр пользовательского компонента.
        /// </summary>
        /// <exception cref="ArgumentException">Возникает в случае, если компонент не отмечен аттрибутом TAttribute.</exception>
        public Component ApplicationComponent {
            get {
                return this._applicationComponent;
            }
            private set {

                #region Проверяем наличие аттрибута TAttribute

                if(value.GetType().GetCustomAttributes(typeof(TAttribute), false).Length==0) {
                    throw new ArgumentException(string.Format("Класс \"{0}\" не отмечен аттрибутом \"{1}\".", value.GetType().FullName, typeof(TAttribute).Name));
                }

                #endregion

                this._components.Add(this._applicationComponent=value);

                #region Подписываемся на события, отмеченные аттрибутом ApplicationProcessedAttribute

                foreach(EventInfo _event in value.GetType().GetEvents(BindingFlags.Public|BindingFlags.Instance)) {
                    if(_event.GetCustomAttributes(typeof(ApplicationProcessedAttribute), false).Length>0) {
                        _EventDispatcher.AddEventDispatcher(value, _event, this._DefaultEventHandler);
                    }
                }

                #endregion
            }
        }

        /// <summary>
        /// Возвращает или устанавливает рабочий каталог.
        /// </summary>
        public string WorkingDirectory {
            get;
            set;
        }

        #endregion

        #region Открытые события

        /// <summary>
        /// Возникает при перехвате событий пользовательского компонента.
        /// </summary>
        public event EventHandler<FireEventArgs> FireEvent;

        #endregion

        /// <summary>
        /// Releases the unmanaged resources used by the ComponentContainer and optionally releases the managed resources.
        /// </summary>
        /// <param name="disposing">true to release both managed and unmanaged resources; false to release only unmanaged resources.</param>
        protected override void Dispose(bool disposing) {
                if(disposing&&this._components!=null) {
                    this._components.Dispose();
                }
            base.Dispose(disposing);
        }

        /// <summary>
        /// Handles the ReflectionOnlyAssemblyResolve event of the CurrentDomain control.
        /// </summary>
        /// <param name="sender">The source of the event.</param>
        /// <param name="args">The <see cref="System.ResolveEventArgs"/> instance containing the event data.</param>
        /// <returns></returns>
        private Assembly _ReflectionOnlyAssemblyResolve(object sender, ResolveEventArgs args) {
            return Assembly.ReflectionOnlyLoad(args.Name);
        }

        private string _CopyFile(string fileName) {
            string _result=Path.Combine(this._AppDirectory, fileName);
            WebRequest _request=WebRequest.Create(new Uri(string.Format("{0}/{1}", this.WorkingDirectory.TrimEnd('/'), fileName)));
            WebResponse _response=_request.GetResponse();
            using(Stream _responseStream=_response.GetResponseStream()) {
                using(Stream _stream=File.Create(_result)) {
                    byte[] _buf=new byte[8*1024];
                    while(true) {
                        _stream.Write(_buf, 0, _responseStream.Read(_buf, 0, _buf.Length));
                        if(_stream.Length==_response.ContentLength) {
                            break;
                        }
                    }
                }
            }
            return _result;
        }

        private bool _IsLoad(string assemblyName,out Assembly assembly) {
            foreach(var _assembly in AppDomain.CurrentDomain.GetAssemblies()) {
                if(Path.GetFileName(_assembly.Location).ToLower()==assemblyName.ToLower()) {
                    assembly=_assembly;
                    return true;
                }
            }
            assembly=null;
            return false;
        }

        private string _AppDirectory {
            get {
                if(this._appDirectory==null) {
                    var _asm=this.GetType().Assembly;
                    var _name=_asm.GetName();
                    var _subdir1=string.Format("{0}_{1}", _name.Name, _name.ProcessorArchitecture);
                    var _subdir2=string.Format("{0}_{1}__", _asm.ImageRuntimeVersion, _name.Version);
                    foreach(var _byte in _name.GetPublicKeyToken()) {
                        _subdir2+=_byte.ToString("x2");
                    }
                    this._appDirectory=Path.Combine(Path.Combine(Path.Combine(Path.GetTempPath(), _subdir1), _subdir2), Guid.NewGuid().ToString());
                    Directory.CreateDirectory(this._appDirectory);
                }
                return this._appDirectory;
            }
        }

        /// <summary>
        /// Обработчик по умолчанию для событий пользовательского элемента управления.
        /// </summary>
        /// <param name="sender">Источник события.</param>
        /// <param name="e">Аргументы события.</param>
        private void _DefaultEventHandler(object sender, _DispatchEventArgs e) {
            if(this.FireEvent!=null) {
                this.FireEvent(this, new FireEventArgs(new EventDescriptor(e.EventName, e.EventArgs)));
            }
        }

        /// <summary>
        /// Используется для подписки на события пользовательского элемента управления.
        /// </summary>
        [ComVisible(false)]
        private sealed class _EventDispatcher {
            private string _event_name;
            private EventHandler<_DispatchEventArgs> _handler;

            /// <summary>
            /// Initializes a new instance of the <see cref="_EventDispatcher"/> class.
            /// </summary>
            /// <param name="target">Ссылка на целевой объект.</param>
            /// <param name="eventInfo">Описание события.</param>
            /// <param name="handler">Обработчик.</param>
            private _EventDispatcher(object target, EventInfo eventInfo, EventHandler<_DispatchEventArgs> handler) {
                this._event_name=eventInfo.Name;
                this._handler=handler;
                eventInfo.AddEventHandler(target, new EventHandler(this._Dispatch));
            }

            /// <summary>
            /// Обработчик события.
            /// </summary>
            /// <param name="sender">The source of the event.</param>
            /// <param name="e">The <see cref="System.EventArgs"/> instance containing the event data.</param>
            private void _Dispatch(object sender, EventArgs e) {
                this._handler(this, new _DispatchEventArgs(this._event_name, sender, e));
            }

            /// <summary>
            /// Добавляет обработчик события.
            /// </summary>
            /// <param name="target">Ссылка на целевой объект.</param>
            /// <param name="eventInfo">Описание события.</param>
            /// <param name="handler">Обработчик.</param>
            /// <returns>Instance of the <see cref="_EventDispatcher"/> class.</returns>
            public static _EventDispatcher AddEventDispatcher(object target, EventInfo eventInfo, EventHandler<_DispatchEventArgs> handler) {
                return new _EventDispatcher(target, eventInfo, handler);
            }
        }

        /// <summary>
        /// Класс аргументов диспетчера событий.
        /// </summary>
        [ComVisible(false)]
        private sealed class _DispatchEventArgs:EventArgs {

            /// <summary>
            /// Initializes a new instance of the <see cref="_DispatchEventArgs"/> class.
            /// </summary>
            /// <param name="eventName">Name of the event.</param>
            /// <param name="sender">The sender.</param>
            /// <param name="eventArgs">The <see cref="System.EventArgs"/> instance containing the event data.</param>
            public _DispatchEventArgs(string eventName, object sender, EventArgs eventArgs) {
                this.EventName=eventName;
                this.Sender=sender;
                this.EventArgs=eventArgs;
            }

            /// <summary>
            /// Gets the name of the event.
            /// </summary>
            /// <value>The name of the event.</value>
            public string EventName {
                get;
                private set;
            }

            /// <summary>
            /// Gets the sender.
            /// </summary>
            /// <value>The sender.</value>
            public object Sender {
                get;
                private set;
            }

            /// <summary>
            /// Gets the event args.
            /// </summary>
            /// <value>The event args.</value>
            public EventArgs EventArgs {
                get;
                private set;
            }
        }

        /// <summary>
        /// Класс аргументов события FireEvent.
        /// </summary>
        [ComVisible(false)]
        internal sealed class FireEventArgs:EventArgs {

            public FireEventArgs(EventDescriptor eventId) {
                this.EventId=eventId;
            }

            public EventDescriptor EventId {
                get;
                private set;
            }
        }
    }
}
