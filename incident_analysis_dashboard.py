import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import calendar
import locale

# Aseta suomenkielinen lokalisaatio
try:
    locale.setlocale(locale.LC_TIME, 'fi_FI.UTF-8')
except:
    try:
        locale.setlocale(locale.LC_TIME, 'Finnish_Finland.1252')  # Windows
    except:
        pass  # Käytä oletusta jos ei onnistu

# Sivun konfiguraatio
st.set_page_config(
    page_title="Hälytysten Analyysihallinta",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Suomenkieliset nimet
FINNISH_MONTHS = [
    "Tammikuu", "Helmikuu", "Maaliskuu", "Huhtikuu", "Toukokuu", "Kesäkuu",
    "Heinäkuu", "Elokuu", "Syyskuu", "Lokakuu", "Marraskuu", "Joulukuu"
]

FINNISH_WEEKDAYS = ["Ma", "Ti", "Ke", "To", "Pe", "La", "Su"]
FINNISH_WEEKDAYS_LONG = ["Maanantai", "Tiistai", "Keskiviikko", "Torstai", "Perjantai", "Lauantai", "Sunnuntai"]

def get_worker_count(hour):
    """Laske työntekijämäärä tunnin perusteella"""
    workers = 0
    
    # Yövuoro 19:15-07:15 (2 henkilöä)
    if hour >= 19 or hour < 7:
        workers += 2
    
    # Aamuvuoro 07:00-17:00 (3 henkilöä)
    if hour >= 7 and hour < 17:
        workers += 3
    
    # Iltavuorot (portaittain)
    if hour >= 9 and hour < 19: workers += 1  # 09:15-19:15
    if hour >= 10 and hour < 20: workers += 1  # 10:00-20:00
    if hour >= 11 and hour < 21: workers += 1  # 11:00-21:00
    if hour >= 13 and hour < 23: workers += 1  # 13:00-23:00
    
    return workers

def get_finnish_weekday(date_obj):
    """Palauttaa suomenkielisen viikonpäivän nimen"""
    if pd.isna(date_obj):
        return "Tuntematon"
    weekday_num = date_obj.weekday()  # 0=Maanantai, 6=Sunnuntai
    return FINNISH_WEEKDAYS_LONG[weekday_num]

def get_finnish_month_name(month_num):
    """Palauttaa suomenkielisen kuukauden nimen"""
    return FINNISH_MONTHS[month_num - 1]

def create_calendar_view(daily_stats):
    """Luo kalenterinäkymä päivittäisistä tilastoista"""
    if len(daily_stats) == 0:
        return None
    
    # Muunna päivämäärät datetime-objekteiksi
    daily_stats = daily_stats.copy()
    daily_stats['date_obj'] = pd.to_datetime(daily_stats['date'])
    
    # Määritä kuukausi ja vuosi
    first_date = daily_stats['date_obj'].min()
    last_date = daily_stats['date_obj'].max()
    
    # Jos kaikki data on samalta kuukaudelta
    if first_date.month == last_date.month and first_date.year == last_date.year:
        month = first_date.month
        year = first_date.year
    else:
        # Käytä ensimmäistä kuukautta
        month = first_date.month
        year = first_date.year
    
    # Luo kuukauden kalenteri
    cal = calendar.monthcalendar(year, month)
    
    # Luo HTML-taulukko
    month_name = get_finnish_month_name(month)
    html = f"""
    <div style="margin: 20px 0;">
        <h3 style="text-align: center; margin-bottom: 20px; color: #1f77b4;">
            {month_name} {year}
        </h3>
        <div style="text-align: right; margin-bottom: 10px; font-size: 14px;">
            <span style="color: #666;">Rauhallisin: </span>
            <span style="background-color: #d4edda; padding: 2px 8px; border-radius: 3px; font-weight: bold;">
                {daily_stats.loc[daily_stats['total_incidents'].idxmin(), 'day']:.0f} ({daily_stats['total_incidents'].min():.0f} inc)
            </span>
        </div>
        <table style="width: 100%; border-collapse: collapse; margin: 0 auto; max-width: 900px;">
            <thead>
                <tr style="background-color: #f8f9fa;">
    """
    
    # Viikonpäivien otsikot
    for day_name in FINNISH_WEEKDAYS:
        html += f'<th style="padding: 10px; text-align: center; border: 1px solid #dee2e6; font-weight: bold; color: #495057;">{day_name}</th>'
    
    html += """
                </tr>
            </thead>
            <tbody>
    """
    
    # Kalenterin rivit
    for week in cal:
        html += '<tr>'
        for day in week:
            if day == 0:
                # Tyhjä päivä
                html += '<td style="padding: 15px; border: 1px solid #dee2e6; background-color: #f8f9fa;"></td>'
            else:
                # Etsi päivän data
                day_data = daily_stats[daily_stats['date_obj'].dt.day == day]
                
                if len(day_data) > 0:
                    row = day_data.iloc[0]
                    
                    # Määritä väri tavoitteiden perusteella
                    if row['day_target_met'] and row['night_target_met']:
                        bg_color = "#d4edda"  # Vihreä - molemmat tavoitteet täytetty
                        border_color = "#28a745"
                        status_text = "Molemmat tavoitteet täytetty"
                    elif row['day_target_met'] or row['night_target_met']:
                        bg_color = "#fff3cd"  # Keltainen - yksi tavoite täytetty
                        border_color = "#ffc107"
                        status_text = "Yksi tavoite täytetty"
                    else:
                        bg_color = "#f8d7da"  # Punainen - kumpikaan tavoite ei täytetty
                        border_color = "#dc3545"
                        status_text = "Kumpikaan tavoite ei täytetty"
                    
                    # Lisää valittu päivä -efekti (voidaan laajentaa myöhemmin)
                    if day == first_date.day:
                        border_style = f"3px solid {border_color}"
                    else:
                        border_style = f"1px solid {border_color}"
                    
                    html += f"""
                    <td style="padding: 8px; border: {border_style}; background-color: {bg_color}; vertical-align: top; position: relative;">
                        <div style="font-weight: bold; font-size: 16px; margin-bottom: 5px;">{day}</div>
                        <div style="font-size: 11px; line-height: 1.2;">
                            <div style="color: #2c5aa0; font-weight: bold;">P: {row['day_shift_avg']:.2f}</div>
                            <div style="color: #6f42c1; font-weight: bold;">Y: {row['night_shift_avg']:.2f}</div>
                            <div style="color: #666; margin-top: 2px;">{row['total_incidents']:.0f} inc</div>
                        </div>
                    </td>
                    """
                else:
                    # Ei dataa tälle päivälle
                    html += f"""
                    <td style="padding: 15px; border: 1px solid #dee2e6; background-color: #ffffff; vertical-align: top;">
                        <div style="font-weight: bold; color: #999;">{day}</div>
                        <div style="font-size: 11px; color: #ccc; margin-top: 5px;">Ei dataa</div>
                    </td>
                    """
        html += '</tr>'
    
    html += """
            </tbody>
        </table>
        <div style="margin-top: 15px; font-size: 12px; color: #666;">
            <div style="display: flex; justify-content: center; gap: 20px; flex-wrap: wrap;">
                <span><span style="display: inline-block; width: 12px; height: 12px; background-color: #d4edda; border: 1px solid #28a745; margin-right: 5px;"></span>Molemmat tavoitteet täytetty</span>
                <span><span style="display: inline-block; width: 12px; height: 12px; background-color: #fff3cd; border: 1px solid #ffc107; margin-right: 5px;"></span>Yksi tavoite täytetty</span>
                <span><span style="display: inline-block; width: 12px; height: 12px; background-color: #f8d7da; border: 1px solid #dc3545; margin-right: 5px;"></span>Kumpikaan tavoite ei täytetty</span>
                <span><span style="display: inline-block; width: 12px; height: 12px; background-color: #ffffff; border: 1px solid #dee2e6; margin-right: 5px;"></span>Ei dataa</span>
            </div>
        </div>
    </div>
    """
    
    return html

@st.cache_data
def process_data(df):
    """Käsittele Excel-data analyysiin"""
    try:
        # Tarkista että tarvittavat sarakkeet löytyvät
        required_columns = ['Hour', 'Incidents handled by agent']
        for col in required_columns:
            if col not in df.columns:
                st.error(f"Saraketta '{col}' ei löydy datasta. Tarkista Excel-tiedosto.")
                return None
        
        # Tee kopio datasta
        df_clean = df.copy()
        
        # Muunna Hour-sarake numeroiksi
        try:
            df_clean['Hour'] = pd.to_numeric(df_clean['Hour'], errors='coerce')
        except Exception as e:
            st.error(f"Virhe Hour-sarakkeen muunnossa: {str(e)}")
            return None
        
        # Muunna Incidents-sarake numeroiksi
        try:
            df_clean['Incidents handled by agent'] = pd.to_numeric(df_clean['Incidents handled by agent'], errors='coerce')
        except Exception as e:
            st.error(f"Virhe Incidents-sarakkeen muunnossa: {str(e)}")
            return None
        
        # Suodata vain validi data
        df_clean = df_clean[
            (df_clean['Hour'].notna()) & 
            (df_clean['Incidents handled by agent'].notna()) &
            (df_clean['Hour'] >= 0) & 
            (df_clean['Hour'] <= 23)
        ].copy()
        
        if len(df_clean) == 0:
            st.error("Ei validia dataa löydetty. Tarkista että Hour-sarake sisältää numeroita 0-23 ja Incidents-sarake sisältää numeroita.")
            return None
        
        # Lisää työntekijämäärät ja laskelmat
        df_clean['workers'] = df_clean['Hour'].apply(get_worker_count)
        df_clean['incidents_per_worker'] = df_clean['Incidents handled by agent'] / df_clean['workers']
        
        # Käsittele päivämäärät
        if 'Date' in df.columns:
            try:
                # Kokeile eri päivämäärämuotoja
                if df_clean['Date'].dtype in ['int64', 'float64']:
                    # Excel serial date
                    df_clean['date'] = pd.to_datetime('1900-01-01') + pd.to_timedelta(df_clean['Date'] - 2, unit='D')
                else:
                    # Tavallinen päivämäärämerkkijono
                    df_clean['date'] = pd.to_datetime(df_clean['Date'], errors='coerce')
                
                # Jos päivämäärien muunto epäonnistui, käytä nykyistä päivää
                if df_clean['date'].isna().all():
                    df_clean['date'] = datetime.now().date()
                    df_clean['date_str'] = df_clean['date'].astype(str)
                    df_clean['day_name'] = 'Tuntematon'
                    df_clean['day'] = 1
                else:
                    df_clean['date_str'] = df_clean['date'].dt.strftime('%Y-%m-%d')
                    df_clean['day_name'] = df_clean['date'].apply(get_finnish_weekday)
                    df_clean['day'] = df_clean['date'].dt.day
                    
            except Exception as e:
                st.warning(f"Päivämäärien käsittely epäonnistui: {str(e)}. Käytetään oletuspäivämääriä.")
                df_clean['date'] = datetime.now().date()
                df_clean['date_str'] = df_clean['date'].astype(str)
                df_clean['day_name'] = 'Tuntematon'
                df_clean['day'] = 1
        else:
            # Jos ei päivämääriä, luo dummy-päivämäärät
            df_clean['date'] = datetime.now().date()
            df_clean['date_str'] = df_clean['date'].astype(str)
            df_clean['day_name'] = 'Tuntematon'
            df_clean['day'] = 1
        
        return df_clean
        
    except Exception as e:
        st.error(f"Virhe datan käsittelyssä: {str(e)}")
        return None

def calculate_hourly_stats(df):
    """Laske tuntikohtaiset tilastot"""
    hourly_stats = []
    
    for hour in range(24):
        hour_data = df[df['Hour'] == hour]
        if len(hour_data) > 0:
            avg_incidents = hour_data['Incidents handled by agent'].mean()
            worker_count = get_worker_count(hour)
            avg_incidents_per_worker = avg_incidents / worker_count
            
            hourly_stats.append({
                'hour': hour,
                'hour_str': f"{hour:02d}:00",
                'avg_incidents': round(avg_incidents, 2),
                'worker_count': worker_count,
                'incidents_per_worker': round(avg_incidents_per_worker, 2),
                'days_count': len(hour_data)
            })
    
    return pd.DataFrame(hourly_stats)

def calculate_daily_stats(df):
    """Laske päivittäiset tilastot"""
    daily_stats = []
    
    for date_str in df['date_str'].unique():
        day_data = df[df['date_str'] == date_str]
        
        # Jaa päivä- ja yötyöntekijöihin
        day_shift = day_data[(day_data['Hour'] >= 7) & (day_data['Hour'] < 23)]
        night_shift = day_data[(day_data['Hour'] >= 23) | (day_data['Hour'] < 7)]
        
        day_shift_avg = day_shift['incidents_per_worker'].mean() if len(day_shift) > 0 else 0
        night_shift_avg = night_shift['incidents_per_worker'].mean() if len(night_shift) > 0 else 0
        
        daily_stats.append({
            'date': date_str,
            'day_name': day_data['day_name'].iloc[0] if len(day_data) > 0 else 'Tuntematon',
            'day': day_data['day'].iloc[0] if len(day_data) > 0 else 1,
            'total_incidents': day_data['Incidents handled by agent'].sum(),
            'day_shift_avg': round(day_shift_avg, 2),
            'night_shift_avg': round(night_shift_avg, 2),
            'day_target_met': day_shift_avg >= 5.1,
            'night_target_met': night_shift_avg >= 4.6
        })
    
    return pd.DataFrame(daily_stats)

def create_combined_chart(hourly_df):
    """Luo yhdistetty kaavio"""
    fig = make_subplots(
        rows=1, cols=1,
        specs=[[{"secondary_y": True}]],
        subplot_titles=["Yhdistetty analyysi"]
    )
    
    # Pylväskaavio incidenteille
    fig.add_trace(
        go.Bar(
            x=hourly_df['hour_str'],
            y=hourly_df['avg_incidents'],
            name='Keskimääräiset incidentit',
            marker_color='lightblue'
        ),
        secondary_y=False
    )
    
    # Viivakaavio incidenteille per työntekijä
    fig.add_trace(
        go.Scatter(
            x=hourly_df['hour_str'],
            y=hourly_df['incidents_per_worker'],
            name='Incidentit/työntekijä',
            line=dict(color='red', width=3),
            mode='lines+markers'
        ),
        secondary_y=True
    )
    
    # Viivakaavio työntekijämäärille
    fig.add_trace(
        go.Scatter(
            x=hourly_df['hour_str'],
            y=hourly_df['worker_count'],
            name='Työntekijämäärä',
            line=dict(color='green', width=2),
            mode='lines+markers'
        ),
        secondary_y=True
    )
    
    fig.update_xaxes(title_text="Kelloaika")
    fig.update_yaxes(title_text="Incidentit", secondary_y=False)
    fig.update_yaxes(title_text="Incidentit/työntekijä & Työntekijämäärä", secondary_y=True)
    
    fig.update_layout(height=500, showlegend=True)
    
    return fig

def main():
    # Otsikko
    st.title("📊 Hälytysten Analyysihallinta")
    st.markdown("**Lataa Excel-tiedosto ja saa automaattinen analyysi hälytysten määrästä suhteessa työntekijöihin**")
    
    # Sivupalkki
    with st.sidebar:
        st.header("⚙️ Asetukset")
        
        # Tiedoston lataus
        uploaded_file = st.file_uploader(
            "Lataa Excel-tiedosto",
            type=['xlsx', 'xls'],
            help="Tiedoston tulee sisältää sarakkeet: 'Hour', 'Incidents handled by agent', ja mahdollisesti 'Date'"
        )
        
        st.markdown("---")
        
        # Vuorojen selitys
        st.subheader("🕐 Vuorojärjestely")
        st.markdown("""
        **Yövuoro** (19:15-07:15): 2 henkilöä  
        **Aamuvuoro** (07:00-17:00): 3 henkilöä  
        **Iltavuorot** (portaittain): 1-4 henkilöä
        - 09:15-19:15: +1 henkilö
        - 10:00-20:00: +1 henkilö  
        - 11:00-21:00: +1 henkilö
        - 13:00-23:00: +1 henkilö
        """)
        
        st.markdown("---")
        
        # Tuottavuustavoitteet
        st.subheader("🎯 Tuottavuustavoitteet")
        st.markdown("""
        **Päivätyöntekijät** (07-23): ≥5.1 inc/työnt./h  
        **Yötyöntekijät** (23-07): ≥4.6 inc/työnt./h
        """)
    
    # Pääsisältö
    if uploaded_file is not None:
        try:
            # Lue Excel-tiedosto
            df = pd.read_excel(uploaded_file)
            st.success(f"✅ Tiedosto ladattu! Löydettiin {len(df)} riviä dataa.")
            
            # Näytä datan otsikko
            with st.expander("📋 Näytä raakadata (ensimmäiset 10 riviä)"):
                st.dataframe(df.head(10))
                
                # Näytä sarakkeiden tietotyypit
                st.subheader("Sarakkeiden tietotyypit:")
                for col in df.columns:
                    st.write(f"- **{col}**: {df[col].dtype}")
            
            # Käsittele data
            processed_df = process_data(df)
            
            if processed_df is not None:
                st.success(f"✅ Data käsitelty onnistuneesti! {len(processed_df)} validia riviä.")
                
                # Laske tilastot
                hourly_stats = calculate_hourly_stats(processed_df)
                daily_stats = calculate_daily_stats(processed_df)
                
                # Tuottavuustavoitteiden analyysi
                day_shift_data = processed_df[(processed_df['Hour'] >= 7) & (processed_df['Hour'] < 23)]
                night_shift_data = processed_df[(processed_df['Hour'] >= 23) | (processed_df['Hour'] < 7)]
                
                day_avg = day_shift_data['incidents_per_worker'].mean() if len(day_shift_data) > 0 else 0
                night_avg = night_shift_data['incidents_per_worker'].mean() if len(night_shift_data) > 0 else 0
                
                # Tulosten näyttäminen
                st.header("🎯 Tuottavuustavoitteiden tulokset")
                
                col1, col2 = st.columns(2)
                
                with col1:
                    day_status = "✅ SAAVUTETTU" if day_avg >= 5.1 else "❌ EI SAAVUTETTU"
                    day_color = "green" if day_avg >= 5.1 else "red"
                    st.markdown(f"""
                    <div style="padding: 20px; border: 2px solid {day_color}; border-radius: 10px; background-color: {'lightgreen' if day_avg >= 5.1 else 'lightcoral'};">
                        <h3>🌅 Päivätyöntekijät (07-23)</h3>
                        <p><strong>Keskiarvo:</strong> {day_avg:.2f} inc/työnt./h</p>
                        <p><strong>Tavoite:</strong> ≥5.1 inc/työnt./h</p>
                        <p><strong>Tulos:</strong> {day_status}</p>
                        <p><strong>Ero:</strong> {day_avg - 5.1:+.2f}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                with col2:
                    night_status = "✅ SAAVUTETTU" if night_avg >= 4.6 else "❌ EI SAAVUTETTU"
                    night_color = "green" if night_avg >= 4.6 else "red"
                    st.markdown(f"""
                    <div style="padding: 20px; border: 2px solid {night_color}; border-radius: 10px; background-color: {'lightgreen' if night_avg >= 4.6 else 'lightcoral'};">
                        <h3>🌙 Yötyöntekijät (23-07)</h3>
                        <p><strong>Keskiarvo:</strong> {night_avg:.2f} inc/työnt./h</p>
                        <p><strong>Tavoite:</strong> ≥4.6 inc/työnt./h</p>
                        <p><strong>Tulos:</strong> {night_status}</p>
                        <p><strong>Ero:</strong> {night_avg - 4.6:+.2f}</p>
                    </div>
                    """, unsafe_allow_html=True)
                
                # Välilehdet eri näkymille
                tab1, tab2, tab3, tab4, tab5 = st.tabs([
                    "📊 Yhdistetty näkymä", 
                    "📈 Tuntikohtainen analyysi", 
                    "📅 Kuukausinäkymä",
                    "📋 Yksityiskohtaiset tilastot",
                    "💡 Suositukset"
                ])
                
                with tab1:
                    st.subheader("Yhdistetty analyysi")
                    if len(hourly_stats) > 0:
                        try:
                            fig_combined = create_combined_chart(hourly_stats)
                            st.plotly_chart(fig_combined, use_container_width=True)
                        except Exception as e:
                            st.error(f"Virhe kaavion luonnissa: {str(e)}")
                            st.info("Näytetään data taulukkona:")
                            st.dataframe(hourly_stats)
                    else:
                        st.warning("Ei dataa kaavion piirtämiseen.")
                
                with tab3:
                    st.subheader("📅 Kuukausinäkymä")
                    
                    if len(daily_stats) > 1:
                        # Luo kalenterinäkymä
                        calendar_html = create_calendar_view(daily_stats)
                        if calendar_html:
                            st.markdown(calendar_html, unsafe_allow_html=True)
                        
                        # Kuukausistatistiikat
                        st.subheader("📊 Kuukauden yhteenveto")
                        col1, col2, col3, col4 = st.columns(4)
                        
                        day_target_met = len(daily_stats[daily_stats['day_target_met']]) 
                        night_target_met = len(daily_stats[daily_stats['night_target_met']])
                        total_days = len(daily_stats)
                        
                        with col1:
                            st.metric("Päivätyöntekijät", f"{day_target_met}/{total_days}", f"{day_target_met/total_days*100:.1f}%")
                        with col2:
                            st.metric("Yötyöntekijät", f"{night_target_met}/{total_days}", f"{night_target_met/total_days*100:.1f}%")
                        with col3:
                            max_day = daily_stats.loc[daily_stats['total_incidents'].idxmax()]
                            st.metric("Kiireisin päivä", f"{max_day['day']:.0f}. ({max_day['day_name']})", f"{max_day['total_incidents']:.0f} inc")
                        with col4:
                            min_day = daily_stats.loc[daily_stats['total_incidents'].idxmin()]
                            st.metric("Rauhallisin päivä", f"{min_day['day']:.0f}. ({min_day['day_name']})", f"{min_day['total_incidents']:.0f} inc")
                        
                        # Päivittäinen kehitys
                        try:
                            fig_daily = px.line(
                                daily_stats, 
                                x='date', 
                                y=['day_shift_avg', 'night_shift_avg'],
                                title='Päivittäinen kehitys',
                                labels={
                                    'value': 'Inc/työnt./h', 
                                    'variable': 'Vuoro',
                                    'date': 'Päivämäärä'
                                }
                            )
                            
                            # Muuta legendan nimet suomeksi
                            fig_daily.for_each_trace(lambda t: t.update(name='Päivätyöntekijät' if 'day_shift_avg' in t.name else 'Yötyöntekijät'))
                            
                            fig_daily.add_hline(y=5.1, line_dash="dash", line_color="red", 
                                              annotation_text="Päivätyöntekijöiden tavoite (5.1)")
                            fig_daily.add_hline(y=4.6, line_dash="dash", line_color="blue", 
                                              annotation_text="Yötyöntekijöiden tavoite (4.6)")
                            st.plotly_chart(fig_daily, use_container_width=True)
                        except Exception as e:
                            st.error(f"Virhe päivittäisen kehityksen kaavion luonnissa: {str(e)}")
                            st.info("Näytetään data taulukkona:")
                            st.dataframe(daily_stats[['date', 'day_shift_avg', 'night_shift_avg']])
                        
                        # Päivittäinen taulukko
                        st.subheader("📋 Päivittäiset tulokset")
                        daily_display = daily_stats.copy()
                        daily_display['Päivätyöntekijät'] = daily_display.apply(
                            lambda x: f"{x['day_shift_avg']:.2f} {'✅' if x['day_target_met'] else '❌'}", axis=1
                        )
                        daily_display['Yötyöntekijät'] = daily_display.apply(
                            lambda x: f"{x['night_shift_avg']:.2f} {'✅' if x['night_target_met'] else '❌'}", axis=1
                        )
                        
                        st.dataframe(
                            daily_display[['date', 'day_name', 'total_incidents', 'Päivätyöntekijät', 'Yötyöntekijät']],
                            column_config={
                                'date': 'Päivämäärä',
                                'day_name': 'Viikonpäivä',
                                'total_incidents': 'Yhteensä inc.',
                                'Päivätyöntekijät': 'Päivätyöntekijät',
                                'Yötyöntekijät': 'Yötyöntekijät'
                            },
                            use_container_width=True
                        )
                    else:
                        st.info("Kuukausinäkymä vaatii useamman päivän dataa.")
                
                with tab4:
                    st.subheader("Tuntikohtaiset tilastot")
                    if len(hourly_stats) > 0:
                        st.dataframe(
                            hourly_stats,
                            column_config={
                                'hour_str': 'Kelloaika',
                                'avg_incidents': 'Keskim. incidentit',
                                'worker_count': 'Työntekijämäärä',
                                'incidents_per_worker': 'Inc/työnt./h',
                                'days_count': 'Päivien lukumäärä'
                            },
                            use_container_width=True
                        )
                    else:
                        st.warning("Ei tilastoja näytettäväksi.")
                
                with tab5:
                    st.subheader("💡 Optimointisuositukset")
                    
                    if len(hourly_stats) > 0:
                        # Ongelmatunnit päivätyöntekijöille
                        day_problems = hourly_stats[
                            (hourly_stats['hour'] >= 7) & 
                            (hourly_stats['hour'] < 23) & 
                            (hourly_stats['incidents_per_worker'] < 5.1)
                        ]
                        
                        # Ongelmatunnit yötyöntekijöille  
                        night_problems = hourly_stats[
                            ((hourly_stats['hour'] >= 23) | (hourly_stats['hour'] < 7)) & 
                            (hourly_stats['incidents_per_worker'] < 4.6)
                        ]
                        
                        col1, col2 = st.columns(2)
                        
                        with col1:
                            st.markdown("### 🌅 Päivätyöntekijät")
                            if len(day_problems) > 0:
                                st.error(f"Ongelmia {len(day_problems)} tunnissa:")
                                for _, row in day_problems.iterrows():
                                    st.write(f"- {row['hour_str']}: {row['incidents_per_worker']} inc/työnt./h")
                                st.markdown("**Suositus:** Vähennä henkilöstöä ali-tuottavina aikoina tai siirrä tehtäviä.")
                            else:
                                st.success("✅ Kaikki tunnit täyttävät tavoitteen!")
                        
                        with col2:
                            st.markdown("### 🌙 Yötyöntekijät")
                            if len(night_problems) > 0:
                                st.error(f"Ongelmia {len(night_problems)} tunnissa:")
                                for _, row in night_problems.iterrows():
                                    st.write(f"- {row['hour_str']}: {row['incidents_per_worker']} inc/työnt./h")
                                st.markdown("**Suositus:** Lisää henkilöstöä ongelmallisina aikoina.")
                            else:
                                st.success("✅ Kaikki tunnit täyttävät tavoitteen!")
                        
                        # Kokonaiskuva
                        st.markdown("### 📊 Kokonaisarvio")
                        if day_avg >= 5.1 and night_avg >= 4.6:
                            st.success("🎉 Molemmat tuottavuustavoitteet saavutettu! Jatka samalla strategialla.")
                        elif day_avg >= 5.1:
                            st.warning("⚠️ Päivätyöntekijöiden tavoite saavutettu, mutta yötyöntekijät tarvitsevat parannusta.")
                        elif night_avg >= 4.6:
                            st.warning("⚠️ Yötyöntekijöiden tavoite saavutettu, mutta päivätyöntekijät tarvitsevat parannusta.")
                        else:
                            st.error("❌ Kumpikaan tuottavuustavoite ei täyty. Tarvitaan merkittäviä toimenpiteitä.")
                    else:
                        st.warning("Ei dataa suositusten tekemiseen.")
        
        except Exception as e:
            st.error(f"Virhe tiedoston käsittelyssä: {str(e)}")
            st.info("Tarkista että Excel-tiedosto sisältää sarakkeet 'Hour' ja 'Incidents handled by agent' ja että ne sisältävät numeroita.")
    
    else:
        # Ohjeet kun ei tiedostoa ladattu
        st.info("👆 Lataa Excel-tiedosto sivupalkista aloittaaksesi analyysin.")
        
        st.markdown("---")
        st.subheader("📋 Käyttöohjeet")
        st.markdown("""
        1. **Lataa Excel-tiedosto** sivupalkista
        2. Tiedoston tulee sisältää vähintään sarakkeet:
           - `Hour` (0-23, numeroina)
           - `Incidents handled by agent` (määrä, numeroina)
           - `Date` (valinnainen, päivämäärille)
        3. **Tarkastele tuloksia** eri välilehdiltä:
           - 📊 Yhdistetty näkymä
           - 📈 Tuntikohtainen analyysi  
           - 📅 Kuukausinäkymä
           - 📋 Tilastot
           - 💡 Suositukset
        """)
        
        st.markdown("### 🎯 Mitä työkalu analysoi:")
        st.markdown("""
        - **Tuottavuustavoitteiden täyttyminen** vuoroittain
        - **Tuntikohtaiset kuormitukset** ja henkilöstötarpeet
        - **Päivittäiset suorituskykytrendit** 
        - **Optimointisuositukset** resurssien allokointiin
        - **Interaktiiviset visualisoinnit** helposti ymmärrettävässä muodossa
        """)

        # Näytä esimerkki oikeasta datamuodosta
        st.markdown("### 📝 Esimerkki oikeasta datamuodosta:")
        example_data = pd.DataFrame({
            'Date': ['2025-02-01', '2025-02-01', '2025-02-01'],
            'Hour': [0, 1, 2],
            'Incidents handled by agent': [9, 14, 16]
        })
        st.dataframe(example_data, use_container_width=True)

if __name__ == "__main__":
    main()lukkona:")
                            st.dataframe(hourly_stats)
                    else:
                        st.warning("Ei dataa kaavion piirtämiseen.")
                
                with tab2:
                    st.subheader("Tuntikohtainen analyysi")
                    
                    if len(hourly_stats) > 0:
                        # Valitse näkymä
                        chart_type = st.selectbox(
                            "Valitse näkymä:",
                            ["Incidentit/työntekijä", "Kokonaisincidentit", "Työntekijämäärät"]
                        )
                        
                        try:
                            if chart_type == "Incidentit/työntekijä":
                                fig = px.line(
                                    hourly_stats, 
                                    x='hour_str', 
                                    y='incidents_per_worker',
                                    title='Incidentit per työntekijä tunnissa',
                                    markers=True
                                )
                                fig.add_hline(y=5.1, line_dash="dash", line_color="red", 
                                             annotation_text="Päivätyöntekijöiden tavoite (5.1)")
                                fig.add_hline(y=4.6, line_dash="dash", line_color="blue", 
                                             annotation_text="Yötyöntekijöiden tavoite (4.6)")
                            
                            elif chart_type == "Kokonaisincidentit":
                                fig = px.bar(
                                    hourly_stats, 
                                    x='hour_str', 
                                    y='avg_incidents',
                                    title='Keskimääräiset incidentit tunneittain'
                                )
                            
                            else:  # Työntekijämäärät
                                fig = px.bar(
                                    hourly_stats, 
                                    x='hour_str', 
                                    y='worker_count',
                                    title='Työntekijämäärät tunneittain'
                                )
                            
                            fig.update_layout(height=500)
                            st.plotly_chart(fig, use_container_width=True)
                            
                        except Exception as e:
                            st.error(f"Virhe kaavion luonnissa: {str(e)}")
                            st.info("Näytetään data tau
