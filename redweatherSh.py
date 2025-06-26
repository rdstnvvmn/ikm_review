import pandas as pd
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from tkinter import messagebox
import tkinter as tk
import calendar

class WeatherApp:
    """Главный класс приложения для анализа погодных данных.

    Attributes:
        data (DataFrame): Загруженные и обработанные погодные данные.
        root (Tk): Главное окно приложения.
        color (str): Основной цвет интерфейса.
        date (datetime): Выбранная пользователем дата для анализа.
        next_date (datetime): Ближайшая доступная дата в данных.
    """

    def __init__(self, root):
        """Инициализирует приложение с главным окном.

        Args:
            root (Tk): Главное окно Tkinter.
        """
        self.data = self._clean_data()
        self.root = root
        self.root.resizable(False, False)
        self.color = '#62639b'
        self._setup_main_window()

    def _clean_data(self):
        """Загружает и подготавливает данные из файла.

        Returns:
            DataFrame: Очищенные и подготовленные данные о погоде.

        Note:
            Переименовывает столбец времени и преобразует его в datetime.
        """
        data = pd.read_excel('weather.xls', skiprows=6)
        data.rename(
            columns={
                'Местное время в Шереметьево / им. А. С. Пушкина (аэропорт)':
                    'Time'
            },
            inplace=True
        )
        data['Time'] = pd.to_datetime(data['Time'], format='%d.%m.%Y %H:%M')
        return data

    def _setup_main_window(self):
        """Настраивает интерфейс главного окна ввода данных.

        Создает:
            - Выпадающие списки для месяца и года
            - Поля ввода для дня, часов и минут
            - Кнопки подтверждения и закрытия
        """
        self.root.title('Погода в Шереметьево')

        months = [f"{i:02d}" for i in range(1, 13)]
        month_label = tk.Label(self.root, text='Месяц')
        month_label.grid(row=0, column=1)
        self.months_var = tk.StringVar()
        months_menu = tk.OptionMenu(self.root, self.months_var, *months)
        months_menu.configure(bg=self.color, width=3)
        months_menu['menu'].configure(bg=self.color)
        months_menu.grid(row=1, column=1, padx=3)

        years = [str(i) for i in range(2005, 2025)]
        year_label = tk.Label(self.root, text='Год', width=3)
        year_label.grid(row=0, column=2)
        self.years_var = tk.StringVar()
        years_menu = tk.OptionMenu(self.root, self.years_var, *years)
        years_menu.configure(bg=self.color, width=3)
        years_menu['menu'].configure(bg=self.color)
        years_menu.grid(row=1, column=2, padx=3)

        day_label = tk.Label(self.root, text='День')
        day_label.grid(row=0, column=0)
        self.day_entry = tk.Entry(self.root, width=5, bg=self.color)
        self.day_entry.grid(row=1, column=0, padx=3)

        hour_label = tk.Label(self.root, text='Час:Минуты')
        hour_label.grid(row=0, column=3, columnspan=3)
        self.hour_entry = tk.Entry(self.root, width=3, bg=self.color)
        self.hour_entry.grid(row=1, column=3)
        colon_label = tk.Label(self.root, text=':')
        colon_label.grid(row=1, column=4)
        self.minute_entry = tk.Entry(self.root, width=3, bg=self.color)
        self.minute_entry.grid(row=1, column=5)

        submit_btn = tk.Button(
            self.root,
            text='Ввод',
            command=self._validate_input,
            bg=self.color,
            width=5
        )
        submit_btn.grid(row=1, column=6, padx=3)

        close_btn = tk.Button(
            self.root,
            text='Закрыть',
            command=self.root.quit,
            bg=self.color
        )
        close_btn.grid(row=1, column=7, padx=3)

    def _validate_input(self):
        """Проверяет корректность введенных пользователем данных.

        Raises:
            ValueError: Если какие-либо данные некорректны или отсутствуют.
        """
        try:
            month = self._get_validated_month()
            year = self._get_validated_year()
            day = self._get_validated_day(year, month)
            hour = self._get_validated_hour()
            minute = self._get_validated_minute()

            self._check_data_exists(year, month, day, hour, minute)
            self._show_results_window(year, month, day, hour, minute)

        except ValueError as e:
            messagebox.showerror('Ошибка', f"Некорректные данные: {e}")

    def _get_validated_month(self):
        """Получает и проверяет введенный месяц.

        Returns:
            int: Номер месяца (1-12).

        Raises:
            ValueError: Если месяц не выбран.
        """
        month = self.months_var.get()
        if not month:
            raise ValueError("Месяц не выбран")
        return int(month)

    def _get_validated_year(self):
        """Получает и проверяет введенный год.

        Returns:
            int: Год (2005-2024).

        Raises:
            ValueError: Если год не выбран.
        """
        year = self.years_var.get()
        if not year:
            raise ValueError("Год не выбран")
        return int(year)

    def _get_validated_day(self, year, month):
        """Получает и проверяет введенный день.

        Args:
            year (int): Год для проверки количества дней в месяце.
            month (int): Месяц для проверки количества дней.

        Returns:
            int: День месяца.

        Raises:
            ValueError: Если день не введен, содержит более 2 цифр
                      или выходит за границы для данного месяца.
        """
        day_str = self.day_entry.get()
        if not day_str:
            raise ValueError("День не введен")
        if len(day_str) > 2:
            raise ValueError("День должен содержать не более 2 цифр")

        day = int(day_str)
        max_days = calendar.monthrange(year, month)[1]
        if day < 1 or day > max_days:
            raise ValueError(
                f"При месяце {month:02d} день должен быть от 1 до {max_days}"
            )
        return day

    def _get_validated_hour(self):
        """Получает и проверяет введенный час.

        Returns:
            int: Час (0-23).

        Raises:
            ValueError: Если час не введен, содержит более 2 цифр
                      или выходит за допустимые границы.
        """
        hour_str = self.hour_entry.get()
        if not hour_str:
            raise ValueError("Час не введен")
        if len(hour_str) > 2:
            raise ValueError("Час должен содержать не более 2 цифр")

        hour = int(hour_str)
        if hour < 0 or hour > 23:
            raise ValueError("Час должен быть от 0 до 23")
        return hour

    def _get_validated_minute(self):
        """Получает и проверяет введенные минуты.

        Returns:
            int: Минуты (0-59).

        Raises:
            ValueError: Если минуты не введены, содержат более 2 цифр
                      или выходят за допустимые границы.
        """
        minute_str = self.minute_entry.get()
        if not minute_str:
            raise ValueError("Минуты не введены")
        if len(minute_str) > 2:
            raise ValueError("Минуты должны содержать не более 2 цифр")

        minute = int(minute_str)
        if minute < 0 or minute > 59:
            raise ValueError("Минуты должны быть от 0 до 59")
        return minute

    def _check_data_exists(self, year, month, day, hour, minute):
        """Проверяет наличие данных для указанной даты.

        Args:
            year (int): Год.
            month (int): Месяц.
            day (int): День.
            hour (int): Час.
            minute (int): Минуты.

        Raises:
            ValueError: Если данные для указанной даты отсутствуют.
        """
        date_str = f"{day:02d}.{month:02d}.{year} {hour:02d}:{minute:02d}"
        self.date = pd.to_datetime(date_str, format='%d.%m.%Y %H:%M')

        filtered_data = self.data[
            (self.data['Time'].dt.year == year) &
            (self.data['Time'].dt.month == month)
            ]

        if filtered_data.empty:
            raise ValueError("Информация на данный месяц отсутствует")

        next_data = self.data[self.data['Time'] >= self.date]
        if not (self.date in self.data['Time'].values):
            if next_data.empty:
                raise ValueError("Приложение не прогноз погоды")
            if len(next_data) == len(self.data):
                raise ValueError("Доступны наблюдения с 1 февраля 2005 года")

        self.next_date = pd.to_datetime(next_data.iloc[-1, 0],
                                        format='%d.%m.%Y %H:%M')

    def _show_results_window(self, year, month, day, hour, minute):
        """Создает и отображает окно с результатами анализа.

        Args:
            year (int): Год.
            month (int): Месяц.
            day (int): День.
            hour (int): Час.
            minute (int): Минуты.
        """
        result_window = tk.Toplevel(self.root)
        result_window.resizable(False, False)
        result_window.title('Результат')
        result_window.geometry('1500x600')

        self._show_weather_info(result_window, year, month, day, hour, minute)
        self._show_avg_temp_graph(result_window, year, month)
        self._show_max_temp_info(result_window, year, month)
        self._show_min_pressure_info(result_window, year, month)
        self._show_day_temp_graph(result_window, year, month, day)

        close_btn = tk.Button(
            result_window,
            text='Закрыть',
            command=result_window.destroy,
            bg=self.color
        )
        close_btn.place(x=750, y=550)

    def _show_weather_info(self, window, year, month, day, hour, minute):
        """Отображает информацию о погоде на указанную дату.

        Args:
            window (Toplevel): Окно для отображения.
            year (int): Год.
            month (int): Месяц.
            day (int): День.
            hour (int): Час.
            minute (int): Минуты.
        """
        weather_text = self._get_weather_text(year, month, day, hour, minute)

        label = tk.Label(
            window,
            text='Информация о погоде на дату: {},'.format(
                self.date.strftime("%Y-%m-%d %H:%M")
            ),
            fg=self.color
        )
        label.place(x=250, y=50, anchor='center')

        result_label = tk.Label(window, text=weather_text)
        result_label.place(x=250, y=75, anchor='center')

    def _get_weather_text(self, year, month, day, hour, minute):
        """Возвращает текстовое описание погоды для указанной даты.

        Args:
            year (int): Год.
            month (int): Месяц.
            day (int): День.
            hour (int): Час.
            minute (int): Минуты.

        Returns:
            str: Описание погоды или сообщение об отсутствии данных.
        """
        if hour in {3, 6, 9, 12, 15, 18, 21}:
            time_str = f"{day:02d}.{month:02d}.{year} {hour:02d}:00"
            exact_date = pd.to_datetime(time_str, format='%d.%m.%Y %H:%M')
            weather = self.data.loc[self.data['Time'] == exact_date, 'WW']

            if not weather.empty and not pd.isna(weather.iloc[0]):
                return weather.iloc[0]

        next_weather = self.data.loc[
            self.data['Time'] == self.next_date,
            'W1'
        ]

        if not next_weather.empty and not pd.isna(next_weather.iloc[0]):
            return next_weather.iloc[0]

        return "Нет информации о погоде"

    def _show_avg_temp_graph(self, window, year, month):
        """Отображает график средней температуры за месяц.

        Args:
            window (Toplevel): Окно для отображения.
            year (int): Год.
            month (int): Месяц.
        """
        label = tk.Label(
            window,
            text='Средняя ежедневная температура за месяц',
            fg=self.color
        )
        label.place(x=1000, y=50, anchor='center')

        fig = self._create_avg_temp_figure(year, month)
        canvas = FigureCanvasTkAgg(fig, master=window)
        canvas.draw()
        canvas.get_tk_widget().place(x=1000, y=175, anchor='center')

    def _create_avg_temp_figure(self, year, month):
        """Создает график средней температуры за месяц.

        Args:
            year (int): Год.
            month (int): Месяц.

        Returns:
            Figure: Объект графика matplotlib.
        """
        filtered = self.data[
            (self.data['Time'].dt.year == year) &
            (self.data['Time'].dt.month == month)
            ]
        daily_avg = filtered.groupby(filtered['Time'].dt.date)['T'].mean()

        fig, ax = plt.subplots(figsize=(9, 2))
        days = daily_avg.index.astype(str).str[8:10]
        ax.plot(days, daily_avg.values, marker='o', color=self.color)
        ax.set_title(f"Средняя температура: {self.date.strftime('%B %Y')}")
        ax.set_xlabel("Дата")
        ax.set_ylabel("Температура (°C)")
        ax.grid(True)
        return fig

    def _show_max_temp_info(self, window, year, month):
        """Отображает информацию о днях с максимальной температурой.

        Args:
            window (Toplevel): Окно для отображения.
            year (int): Год.
            month (int): Месяц.
        """
        max_temp, max_days = self._get_max_temp_info(year, month)

        header = tk.Label(
            window,
            text='Дни с максимальной температурой в месяце',
            fg=self.color
        )
        header.place(x=250, y=200, anchor='center')

        temp_label = tk.Label(
            window,
            text=f'Максимальная температура: {max_temp:.2f}°C'
        )
        temp_label.place(x=250, y=225, anchor='center')

        days_label = tk.Label(
            window,
            text=f'Дни: {",".join(day[-2:] for day in max_days)}'
        )
        days_label.place(x=250, y=250, anchor='center')

    def _get_max_temp_info(self, year, month):
        """Возвращает информацию о максимальной температуре за месяц.

        Args:
            year (int): Год.
            month (int): Месяц.

        Returns:
            tuple: (максимальная температура, список дней с этой температурой)
        """
        filtered = self.data[
            (self.data['Time'].dt.year == year) &
            (self.data['Time'].dt.month == month)
            ]
        daily_avg = filtered.groupby(filtered['Time'].dt.date)['T'].mean()
        max_temp = daily_avg.max()
        max_days = daily_avg[daily_avg == max_temp].index
        max_days_str = [day.strftime('%Y-%m-%d') for day in max_days]
        return max_temp, max_days_str

    def _show_min_pressure_info(self, window, year, month):
        """Отображает информацию о днях с минимальным давлением.

        Args:
            window (Toplevel): Окно для отображения.
            year (int): Год.
            month (int): Месяц.
        """
        min_press, min_days = self._get_min_pressure_info(year, month)

        header = tk.Label(
            window,
            text='Минимальное давление за месяц',
            fg=self.color
        )
        header.place(x=250, y=375, anchor='center')

        press_label = tk.Label(
            window,
            text=f'Минимальное давление: {min_press}'
        )
        press_label.place(x=250, y=400, anchor='center')

        days_label = tk.Label(
            window,
            text=f'Дни: {",".join(str(day.day) for day in min_days)}'
        )
        days_label.place(x=250, y=425, anchor='center')

    def _get_min_pressure_info(self, year, month):
        """Возвращает информацию о минимальном давлении за месяц.

        Args:
            year (int): Год.
            month (int): Месяц.

        Returns:
            tuple: (минимальное давление, список дней с этим давлением)
        """
        filtered = self.data[
            (self.data['Time'].dt.year == year) &
            (self.data['Time'].dt.month == month)
            ]
        min_press = filtered['P'].min()
        min_days = filtered.loc[filtered['P'] == min_press, 'Time']
        return min_press, min_days

    def _show_day_temp_graph(self, window, year, month, day):
        """Отображает график температуры за указанный день.

        Args:
            window (Toplevel): Окно для отображения.
            year (int): Год.
            month (int): Месяц.
            day (int): День.
        """
        label = tk.Label(
            window,
            text='Температура в течение дня',
            fg=self.color
        )
        label.place(x=1000, y=300, anchor='center')

        fig = self._create_day_temp_figure(year, month, day)
        canvas = FigureCanvasTkAgg(fig, master=window)
        canvas.draw()
        canvas.get_tk_widget().place(x=1000, y=425, anchor='center')

    def _create_day_temp_figure(self, year, month, day):
        """Создает график температуры за указанный день.

        Args:
            year (int): Год.
            month (int): Месяц.
            day (int): День.

        Returns:
            Figure: Объект графика matplotlib.
        """
        day_data = self.data[
            (self.data['Time'].dt.year == year) &
            (self.data['Time'].dt.month == month) &
            (self.data['Time'].dt.day == day)
            ]

        fig, ax = plt.subplots(figsize=(9, 2))
        hours = day_data['Time'].dt.hour
        ax.plot(hours, day_data['T'], marker='o', color=self.color)
        ax.set_title(f"Температура {day:02d}.{month:02d}.{year}")
        ax.set_xlabel("Часы")
        ax.set_ylabel("Температура (°C)")
        ax.grid(True)
        return fig


def main():
    """Точка входа в приложение. Создает и запускает главное окно."""
    root = tk.Tk()
    app = WeatherApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()