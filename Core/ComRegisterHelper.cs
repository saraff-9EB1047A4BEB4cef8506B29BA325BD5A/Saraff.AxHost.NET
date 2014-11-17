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
using System.Reflection;
using Microsoft.Win32;

namespace Saraff.AxHost.Core {

    internal static class ComRegisterHelper {
        private const string _controlKey="Control";
        private const string _codebaseKey="CodeBase";
        private const string _inprocKey="InprocServer32";

        /// <summary>
        /// Регистрирует компонент в реестре.
        /// </summary>
        /// <param name="key">Ключ реестра.</param>
        public static void Register(string key) {
            // Open the CLSID\{guid} key for write access
            RegistryKey _key=ComRegisterHelper._OpenSubKey(key);
            try {
                // And create the 'Control' key - this allows it to show up in the ActiveX control container 
                _key.CreateSubKey(ComRegisterHelper._controlKey).Close();

                // Next create the CodeBase entry - needed if not string named and GACced.
                RegistryKey _inprocKey=ComRegisterHelper._OpenInprocServer32(_key);
                try {
                    _inprocKey.SetValue(ComRegisterHelper._codebaseKey, Assembly.GetExecutingAssembly().CodeBase);
                } finally {
                    _inprocKey.Close();
                }
            } finally {
                _key.Close(); // Finally close the main key
            }
        }

        /// <summary>
        /// Удаляет регистрацию компонента из реестра.
        /// </summary>
        /// <param name="key">Ключ реестра.</param>
        public static void Unregister(string key) {
            // Open HKCR\CLSID\{guid} for write access
            RegistryKey _key=ComRegisterHelper._OpenSubKey(key);
            try {
                // Delete the 'Control' key, but don't throw an exception if it does not exist
                _key.DeleteSubKey(ComRegisterHelper._controlKey, false);

                // Next open up InprocServer32
                RegistryKey _inprocKey=ComRegisterHelper._OpenInprocServer32(_key);
                try {
                    // And delete the CodeBase key, again not throwing if missing 
                    _inprocKey.DeleteSubKey(ComRegisterHelper._codebaseKey, false);
                } finally {
                    _inprocKey.Close();
                }
            } finally {
                _key.Close(); // Finally close the main key 
            }
        }

        private static RegistryKey _OpenSubKey(string key) {
            // Open HKCR\CLSID\{guid} for write access
            return Registry.ClassesRoot.OpenSubKey(key.Replace(Registry.ClassesRoot.Name, string.Empty).TrimStart('\\'), true);
        }

        private static RegistryKey _OpenInprocServer32(RegistryKey key) {
            return key.OpenSubKey(ComRegisterHelper._inprocKey, true);
        }
    }
}
