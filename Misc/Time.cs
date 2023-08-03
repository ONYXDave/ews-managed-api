/*
 * Exchange Web Services Managed API
 *
 * Copyright (c) Microsoft Corporation
 * All rights reserved.
 *
 * MIT License
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy of this
 * software and associated documentation files (the "Software"), to deal in the Software
 * without restriction, including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
 * to whom the Software is furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all copies or
 * substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED *AS IS*, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,
 * INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR
 * PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR
 * OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
 * DEALINGS IN THE SOFTWARE.
 */

namespace Microsoft.Exchange.WebServices.Data
    {
    using System;

    /// <summary>
    /// Represents a time.
    /// </summary>
    internal sealed class Time
        {
        private int hours;
        private int minutes;
        private int seconds;

        /// <summary>
        /// Initializes a new instance of Time.
        /// </summary>
        internal Time()
            {
            }

        /// <summary>
        /// Initializes a new instance of Time.
        /// </summary>
        /// <param name="minutes">The number of minutes since 12:00AM.</param>
        internal Time(int minutes)
            : this()
            {
            if (minutes < 0 || minutes >= 1440)
                {
                throw new ArgumentException(
                    Strings.MinutesMustBeBetween0And1439,
                    "minutes");
                }

            Hours = minutes / 60;
            Minutes = minutes % 60;
            Seconds = 0;
            }

        /// <summary>
        /// Initializes a new instance of Time.
        /// </summary>
        /// <param name="dateTime">The DateTime to extract the time part of.</param>
        internal Time(DateTime dateTime)
            {
            Hours = dateTime.Hour;
            Minutes = dateTime.Minute;
            Seconds = dateTime.Second;
            }

        /// <summary>
        /// Initializes a new instance of Time.
        /// </summary>
        /// <param name="hours">The hours.</param>
        /// <param name="minutes">The minutes.</param>
        /// <param name="seconds">The seconds.</param>
        internal Time(int hours, int minutes, int seconds)
            : this()
            {
            Hours = hours;
            Minutes = minutes;
            Seconds = seconds;
            }

        /// <summary>
        /// Convert Time to XML Schema time.
        /// </summary>
        /// <returns>String in XML Schema time format.</returns>
        internal string ToXSTime()
            {
            return string.Format(
                "{0:00}:{1:00}:{2:00}",
                Hours,
                Minutes,
                Seconds);
            }

        /// <summary>
        /// Converts the time into a number of minutes since 12:00AM.
        /// </summary>
        /// <returns>The number of minutes since 12:00AM the time represents.</returns>
        internal int ConvertToMinutes()
            {
            return Minutes + (Hours * 60);
            }

        /// <summary>
        /// Gets or sets the hours.
        /// </summary>
        internal int Hours
            {
            get
                {
                return hours;
                }

            set
                {
                if (value >= 0 && value < 24)
                    {
                    hours = value;
                    }
                else
                    {
                    throw new ArgumentException(Strings.HourMustBeBetween0And23);
                    }
                }
            }

        /// <summary>
        /// Gets or sets the minutes.
        /// </summary>
        internal int Minutes
            {
            get
                {
                return minutes;
                }

            set
                {
                if (value >= 0 && value < 60)
                    {
                    minutes = value;
                    }
                else
                    {
                    throw new ArgumentException(Strings.MinuteMustBeBetween0And59);
                    }
                }
            }

        /// <summary>
        /// Gets or sets the seconds.
        /// </summary>
        internal int Seconds
            {
            get
                {
                return seconds;
                }

            set
                {
                if (value >= 0 && value < 60)
                    {
                    seconds = value;
                    }
                else
                    {
                    throw new ArgumentException(Strings.SecondMustBeBetween0And59);
                    }
                }
            }
        }
    }