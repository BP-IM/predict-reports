import io
import re

import pandas as pd
import streamlit as st


# =========================================================
# PAGE
# =========================================================

st.set_page_config(
    page_title="Predict Reports",
    page_icon="📊",
    layout="wide"
)

st.title("📊 Predict Reports")
st.caption("Отчеттарды автоматты түрде дайындау")


# =========================================================
# COMMON FUNCTIONS
# =========================================================

def find_column(df, keyword):
    keyword = keyword.lower()

    for col in df.columns:
        if keyword in str(col).strip().lower():
            return col

    return None


def make_full_times():
    return pd.date_range(
        "00:00",
        "23:45",
        freq="15min"
    ).strftime("%H:%M")


def make_15min(df, time_col, value_col, value_name):

    work = df[[time_col, value_col]].copy()

    work.columns = [
        "Время",
        value_name
    ]

    work["Время"] = (
        work["Время"]
        .astype(str)
        .str.strip()
    )

    # Тек нақты HH:MM
    work = work[
        work["Время"].str.match(
            r"^([01]\d|2[0-3]):(00|15|30|45)$",
            na=False
        )
    ]

    work[value_name] = pd.to_numeric(
        work[value_name],
        errors="coerce"
    ).fillna(0)

    # Бір уақытта бірнеше жол болса — қосамыз
    grouped = (
        work
        .groupby("Время")[value_name]
        .sum()
    )

    full_times = make_full_times()

    grouped = grouped.reindex(
        full_times,
        fill_value=0
    )

    result = pd.DataFrame({
        "Время": full_times,
        value_name: grouped.values
    })

    result[value_name] = (
        result[value_name]
        .fillna(0)
        .astype(int)
    )

    return result


def make_hourly(df15, value_name):

    df = df15.copy()

    df["Час"] = (
        df["Время"].str[:2]
        + ":00"
    )

    result = (
        df
        .groupby(
            "Час",
            as_index=False
        )[value_name]
        .sum()
    )

    result.columns = [
        "Время",
        value_name
    ]

    result[value_name] = (
        result[value_name]
        .astype(int)
    )

    return result


def create_excel(data_by_date):

    output = io.BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for date_value in sorted(
            data_by_date.keys()
        ):

            df = data_by_date[
                date_value
            ]

            if hasattr(
                date_value,
                "strftime"
            ):
                sheet_name = (
                    date_value.strftime(
                        "%d.%m.%Y"
                    )
                )
            else:
                sheet_name = str(
                    date_value
                )

            # Excel sheet максимум 31 символ
            sheet_name = sheet_name[:31]

            df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

            ws = writer.book[
                sheet_name
            ]

            ws.column_dimensions[
                "A"
            ].width = 15

            ws.column_dimensions[
                "B"
            ].width = 25

            ws.freeze_panes = "A2"

    output.seek(0)

    return output.getvalue()


# =========================================================
# ПОЧАСОВОЙ ТО / ГОСТИ
# =========================================================

def parse_hourly_to(file_bytes):

    excel = pd.ExcelFile(
        io.BytesIO(file_bytes)
    )

    sheet_name = (
        excel.sheet_names[0]
    )

    raw = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=None
    )

    # Header табу
    header_row = None

    for i, row in raw.iterrows():

        text = " | ".join(
            str(x).lower()
            for x in row
            if pd.notna(x)
        )

        if (
            "учетный день" in text
            and "15-минутный интервал" in text
            and "количество гостей" in text
        ):
            header_row = i
            break

    if header_row is None:
        raise ValueError(
            "Почасовой ТО таблицасы табылмады"
        )

    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=sheet_name,
        header=header_row
    )

    date_col = find_column(
        df,
        "учетный день"
    )

    time_col = find_column(
        df,
        "15-минутный интервал"
    )

    guests_col = find_column(
        df,
        "количество гостей"
    )

    if date_col is None:
        raise ValueError(
            "Учетный день бағаны табылмады"
        )

    if time_col is None:
        raise ValueError(
            "15-минутный интервал бағаны табылмады"
        )

    if guests_col is None:
        raise ValueError(
            "Количество гостей бағаны табылмады"
        )

    current_date = None

    rows = []

    for _, row in df.iterrows():

        raw_date = row[
            date_col
        ]

        # Күнді жаңарту
        if pd.notna(raw_date):

            text_date = str(
                raw_date
            ).strip()

            # "Итого", "всего" жолдарын алмау
            if (
                "всего"
                not in text_date.lower()
                and
                "итого"
                not in text_date.lower()
            ):

                parsed_date = (
                    pd.to_datetime(
                        raw_date,
                        dayfirst=True,
                        errors="coerce"
                    )
                )

                if (
                    pd.notna(
                        parsed_date
                    )
                    and
                    parsed_date.year
                    >= 2020
                ):

                    current_date = (
                        parsed_date.date()
                    )

        if current_date is None:
            continue

        raw_time = row[
            time_col
        ]

        if pd.isna(raw_time):
            continue

        time_text = str(
            raw_time
        ).strip()

        if not re.fullmatch(
            r"([01]\d|2[0-3]):(00|15|30|45)",
            time_text
        ):
            continue

        guests = pd.to_numeric(
            row[guests_col],
            errors="coerce"
        )

        if pd.isna(
            guests
        ):
            guests = 0

        rows.append({
            "Дата": current_date,
            "Время": time_text,
            "Количество гостей":
                guests
        })

    result = pd.DataFrame(
        rows
    )

    if result.empty:
        raise ValueError(
            "Гостьтар бойынша дерек табылмады"
        )

    return result


# =========================================================
# PMIX / БЛЮДА
# =========================================================

def parse_pmix(
    file_bytes,
    filename
):

    raw = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=0,
        header=None
    )

    # -----------------------------------------------------
    # Датаны табу
    # -----------------------------------------------------

    report_date = None

    for _, row in raw.iterrows():

        for value in row:

            if pd.isna(value):
                continue

            text = str(
                value
            ).strip()

            match = re.search(
                r"(\d{1,2}\.\d{1,2}\.\d{4})",
                text
            )

            if match:

                parsed = (
                    pd.to_datetime(
                        match.group(1),
                        dayfirst=True,
                        errors="coerce"
                    )
                )

                if pd.notna(
                    parsed
                ):
                    report_date = (
                        parsed.date()
                    )
                    break

        if report_date:
            break

    # Filename fallback
    if report_date is None:

        match = re.search(
            r"(\d{1,2}\.\d{1,2}\.\d{4})",
            filename
        )

        if match:

            report_date = (
                pd.to_datetime(
                    match.group(1),
                    dayfirst=True
                ).date()
            )

    if report_date is None:
        raise ValueError(
            f"{filename}: дата табылмады"
        )

    # -----------------------------------------------------
    # Header
    # -----------------------------------------------------

    header_row = None

    for index, row in raw.iterrows():

        text = " | ".join(
            str(x).lower()
            for x in row
            if pd.notna(x)
        )

        if (
            "15-минутный интервал"
            in text
            and
            "количество блюд"
            in text
        ):

            header_row = (
                index
            )
            break

    if header_row is None:
        raise ValueError(
            f"{filename}: PMIX таблица табылмады"
        )

    df = pd.read_excel(
        io.BytesIO(file_bytes),
        sheet_name=0,
        header=header_row
    )

    time_col = find_column(
        df,
        "15-минутный интервал"
    )

    dishes_col = find_column(
        df,
        "количество блюд"
    )

    if time_col is None:
        raise ValueError(
            f"{filename}: уақыт табылмады"
        )

    if dishes_col is None:
        raise ValueError(
            f"{filename}: Количество блюд табылмады"
        )

    result15 = make_15min(
        df,
        time_col,
        dishes_col,
        "Количество блюд"
    )

    return (
        report_date,
        result15
    )


# =========================================================
# UI TABS
# =========================================================

tab_guests, tab_dishes = (
    st.tabs([
        "👥 Гости",
        "🍽 Проданные блюда"
    ])
)


# =========================================================
# TAB 1 — ГОСТИ
# =========================================================

with tab_guests:

    st.header(
        "Количество гостей"
    )

    st.write(
        "Почасовой ТО файлын сал. "
        "Сайт күндерді өзі бөліп, "
        "15 минут және 1 сағат форматтарын жасайды."
    )

    guest_file = (
        st.file_uploader(
            "Почасовой ТО Excel",
            type=["xlsx"],
            key="guest_file"
        )
    )

    if guest_file:

        try:

            guest_data = (
                parse_hourly_to(
                    guest_file.getvalue()
                )
            )

            dates = sorted(
                guest_data[
                    "Дата"
                ].unique()
            )

            guest_15 = {}
            guest_hour = {}

            info = []

            for date_value in dates:

                day = guest_data[
                    guest_data[
                        "Дата"
                    ]
                    == date_value
                ].copy()

                # 15 минут
                temp = day[
                    [
                        "Время",
                        "Количество гостей"
                    ]
                ].copy()

                full15 = make_15min(
                    temp,
                    "Время",
                    "Количество гостей",
                    "Количество гостей"
                )

                hourly = make_hourly(
                    full15,
                    "Количество гостей"
                )

                guest_15[
                    date_value
                ] = full15

                guest_hour[
                    date_value
                ] = hourly

                real = (
                    day[
                        "Время"
                    ].nunique()
                )

                info.append({
                    "Дата":
                        date_value.strftime(
                            "%d.%m.%Y"
                        ),
                    "Бар интервал":
                        real,
                    "0 қосылды":
                        96 - real
                })

            st.success(
                f"{len(dates)} күн табылды"
            )

            st.dataframe(
                pd.DataFrame(
                    info
                ),
                use_container_width=True,
                hide_index=True
            )

            selected = (
                st.selectbox(
                    "Күнді тексер",
                    dates,
                    format_func=lambda x:
                        x.strftime(
                            "%d.%m.%Y"
                        ),
                    key="guest_date"
                )
            )

            p1, p2 = st.tabs([
                "15 минут",
                "1 сағат"
            ])

            with p1:

                st.dataframe(
                    guest_15[
                        selected
                    ],
                    use_container_width=True,
                    hide_index=True,
                    height=500
                )

            with p2:

                st.dataframe(
                    guest_hour[
                        selected
                    ],
                    use_container_width=True,
                    hide_index=True
                )

            guest_excel15 = (
                create_excel(
                    guest_15
                )
            )

            guest_excel_hour = (
                create_excel(
                    guest_hour
                )
            )

            c1, c2 = st.columns(
                2
            )

            with c1:

                st.download_button(
                    "Скачать гостей — 15 минут",
                    guest_excel15,
                    file_name=(
                        "Гости_15_минут.xlsx"
                    ),
                    mime=(
                        "application/vnd."
                        "openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )

            with c2:

                st.download_button(
                    "Скачать гостей — 1 час",
                    guest_excel_hour,
                    file_name=(
                        "Гости_по_часам.xlsx"
                    ),
                    mime=(
                        "application/vnd."
                        "openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    use_container_width=True
                )

        except Exception as e:

            st.error(
                f"Ошибка: {e}"
            )


# =========================================================
# TAB 2 — БЛЮДА
# =========================================================

with tab_dishes:

    st.header(
        "Проданные блюда"
    )

    st.write(
        "1–7 немесе одан да көп PMIX файл сал. "
        "Әр файлдың датасы автоматты анықталады."
    )

    pmix_files = (
        st.file_uploader(
            "PMIX Excel файлдары",
            type=["xlsx"],
            accept_multiple_files=True,
            key="pmix_files"
        )
    )

    if pmix_files:

        dishes_days = {}

        info = []

        errors = []

        for file in pmix_files:

            try:

                report_date, df15 = (
                    parse_pmix(
                        file.getvalue(),
                        file.name
                    )
                )

                # Бір дата екі рет болса қосу
                if (
                    report_date
                    in dishes_days
                ):

                    combined = (
                        pd.concat([
                            dishes_days[
                                report_date
                            ],
                            df15
                        ])
                    )

                    combined = (
                        combined
                        .groupby(
                            "Время",
                            as_index=False
                        )[
                            "Количество блюд"
                        ]
                        .sum()
                    )

                    dishes_days[
                        report_date
                    ] = combined

                else:

                    dishes_days[
                        report_date
                    ] = df15

                real_intervals = (
                    df15[
                        "Количество блюд"
                    ]
                    .ne(0)
                    .sum()
                )

                info.append({
                    "Дата":
                        report_date.strftime(
                            "%d.%m.%Y"
                        ),
                    "Файл":
                        file.name,
                    "Ненулевых":
                        real_intervals
                })

            except Exception as e:

                errors.append(
                    f"{file.name}: {e}"
                )

        if dishes_days:

            st.success(
                f"{len(pmix_files)} файл жүктелді, "
                f"{len(dishes_days)} күн табылды"
            )

            st.dataframe(
                pd.DataFrame(
                    info
                ),
                use_container_width=True,
                hide_index=True
            )

            dates = sorted(
                dishes_days.keys()
            )

            selected = (
                st.selectbox(
                    "Күнді тексер",
                    dates,
                    format_func=lambda x:
                        x.strftime(
                            "%d.%m.%Y"
                        ),
                    key="dish_date"
                )
            )

            st.dataframe(
                dishes_days[
                    selected
                ],
                use_container_width=True,
                hide_index=True,
                height=500
            )

            dishes_excel = (
                create_excel(
                    dishes_days
                )
            )

            st.download_button(
                "Скачать общий Excel — блюда 15 минут",
                dishes_excel,
                file_name=(
                    "Проданные_блюда_"
                    "15_минут.xlsx"
                ),
                mime=(
                    "application/vnd."
                    "openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="primary",
                use_container_width=True
            )

        if errors:

            st.subheader(
                "Ошибки"
            )

            for error in errors:
                st.error(
                    error
                )