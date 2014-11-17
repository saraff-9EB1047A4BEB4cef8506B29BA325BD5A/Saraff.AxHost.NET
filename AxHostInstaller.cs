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
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Text;


namespace Saraff.AxHost {

    /// <summary>
    /// Выполняет регистрацию сборки в системею
    /// </summary>
    [RunInstaller(true)]
    public partial class AxHostInstaller:System.Configuration.Install.Installer {

        /// <summary>
        /// Initializes a new instance of the <see cref="AxHostInstaller"/> class.
        /// </summary>
        public AxHostInstaller() {
            InitializeComponent();
        }

        protected override void OnAfterInstall(IDictionary savedState) {
            try {
                this._Install();
            } catch(Exception ex) {
                this._WriteLog(string.Format("{1}: {2}{0}{3}{0}", Environment.NewLine, ex.GetType().Name, ex.Message, ex.StackTrace));
            }
            base.OnAfterInstall(savedState);
        }

        protected override void OnAfterUninstall(IDictionary savedState) {
            try {
                this._Uninstall();
            } catch(Exception ex) {
                this._WriteLog(string.Format("{1}: {2}{0}{3}{0}",Environment.NewLine,ex.GetType().Name,ex.Message,ex.StackTrace));
            }
            base.OnAfterUninstall(savedState);
        }

        private void _Install() {
            this._ExecuteRegAsm("/registered");
        }

        private void _Uninstall() {
            this._ExecuteRegAsm("/unregister");
        }

        private void _ExecuteRegAsm(string options) {
            ProcessStartInfo _si=new ProcessStartInfo();
            _si.WorkingDirectory=Path.GetDirectoryName(this.GetType().Assembly.Location);
            _si.Arguments=string.Format("\"{0}\" /tlb /codebase /nologo {1}", this.Context.Parameters["assemblypath"], options);
            _si.FileName=this._RegAsmPath;
            _si.CreateNoWindow=true;
            _si.UseShellExecute=false;
            _si.RedirectStandardOutput=true;
            _si.RedirectStandardError=true;
            using(Process _proc=Process.Start(_si)) {
                _proc.WaitForExit();
                this._WriteLog("-------- RegAsm Output --------");
                this._WriteLog(new StreamReader(_proc.StandardOutput.BaseStream, Encoding.GetEncoding(866)).ReadToEnd());
                this._WriteLog("-------- RegAsm Error ---------");
                this._WriteLog(new StreamReader(_proc.StandardError.BaseStream, Encoding.GetEncoding(866)).ReadToEnd());
                this._WriteLog("------------ End --------------");
            }
        }

        private string _RegAsmPath {
            get {
                Version _ver=Environment.Version;
                return Path.Combine(Environment.SystemDirectory, string.Format("..\\Microsoft.NET\\Framework\\v{0}.{1}.{2}\\RegAsm.exe", _ver.Major, _ver.Minor, _ver.Build));
            }
        }

        private void _WriteLog(string message) {
            Debug.WriteLine(message);
        }
    }
}
