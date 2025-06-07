import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import calendar

# Sivun konfiguraatio
st.set_page_config(
    page_title="Incident Analysis Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

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

def process_data(df):
    """Käsittele Excel-data analyysiin"""
    try:
        # Tarkista että tarvittavat sarakkeet löytyvät
        required_columns = ['Hour', 'Incidents handled by agent']
        for col in required_columns:
            if col not in df.columns:
                st.error(f"Saraketta '{col}' ei löydy datasta. Tarkista Excel-tiedosto.")
                return None
        
        # Suodata vain validi data
        df_clean = df[
            (df['Hour'].notna()) & 
            (df['Incidents handled by agent'].notna()) &
            (df['Hour'] >= 0) & 
            (df['Hour'] <= 23)
        ].copy()
        
        if len(df_clean) == 0:
            st.error("Ei validia dataa löydetty.")
            return None
        
        # Lisää työntekijämäärät ja laskelmat
        df_clean['workers'] = df_clean['Hour'].apply(get_worker_count)
        df_clean['incidents_per_worker'] = df_clean['Incidents handled by agent'] / df_clean['workers']
        
        # Käsittele päivämäärät
        if 'Date' in df.columns:
            # Excel serial date → datetime
            if df_clean['Date'].dtype in ['int64', 'float64']:
                df_clean['date'] = pd.to_datetime('1900-01-01') + pd.to_timedelta(df_clean['Date'] - 2, unit='D')
            else:
                df_clean['date'] = pd.to_datetime(df_clean['Date'])
            
            df_clean['date_str'] = df_clean['date'].dt.strftime('%Y-%m-%d')
            df_clean['day_name'] = df_clean['date'].dt.day_name()
            df_clean['day'] = df_clean['date'].dt.day
        else:
            # Jos ei päivämääriä, luo dummy-päivämäärät
            df_clean['date'] = datetime.now().date()
            df_clean['date_str'] = df_clean['date'].astype(str)
            df_clean['day_name'] = 'Unknown'
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
            'day_name': day_data['day_name'].iloc[0] if len(day_data) > 0 else 'Unknown',
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
        secondary_y=True,
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

def create_monthly_calendar(daily_df):
    """Luo kuukausikalenteri"""
    if len(daily_df) == 0:
        return None
    
    # Ryhmittele kuukausittain
    daily_df['month'] = pd.to_datetime(daily_df['date']).dt.month
    daily_df['year'] = pd.to_datetime(daily_df['date']).dt.year
    
    # Ota ensimmäinen kuukausi
    first_month = daily_df['month'].iloc[0]
    first_year = daily_df['year'].iloc[0]
    month_data = daily_df[(daily_df['month'] == first_month) & (daily_df['year'] == first_year)]
    
    return month_data

def main():
    # Otsikko
    st.title("📊 Incident Analysis Dashboard")
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
            
            # Käsittele data
            processed_df = process_data(df)
            
            if processed_df is not None:
                # Laske tilastot
                hourly_stats = calculate_hourly_stats(processed_df)
                daily_stats = calculate_daily_stats(processed_df)
                
                # Tuottavuustavoitteiden analyysi
                day_shift_data = processed_df[(processed_df['Hour'] >= 7) & (processed_df['Hour'] < 23)]
                night_shift_data = processed_df[(processed_df['Hour'] >= 23) | (processed_df['Hour'] < 7)]
                
                day_avg = day_shift_data['incidents_per_worker'].mean()
                night_avg = night_shift_data['incidents_per_worker'].mean()
                
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
                    fig_combined = create_combined_chart(hourly_stats)
                    st.plotly_chart(fig_combined, use_container_width=True)
                
                with tab2:
                    st.subheader("Tuntikohtainen analyysi")
                    
                    # Valitse näkymä
                    chart_type = st.selectbox(
                        "Valitse näkymä:",
                        ["Incidentit/työntekijä", "Kokonaisincidentit", "Työntekijämäärät"]
                    )
                    
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
                
                with tab3:
                    st.subheader("Kuukausinäkymä")
                    
                    if len(daily_stats) > 1:
                        # Kuukausistatistiikat
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
                            st.metric("Kiireisin päivä", f"{max_day['day']} ({max_day['day_name']})", f"{max_day['total_incidents']} inc")
                        with col4:
                            min_day = daily_stats.loc[daily_stats['total_incidents'].idxmin()]
                            st.metric("Rauhallisinta", f"{min_day['day']} ({min_day['day_name']})", f"{min_day['total_incidents']} inc")
                        
                        # Päivittäinen kehitys
                        fig_daily = px.line(
                            daily_stats, 
                            x='date', 
                            y=['day_shift_avg', 'night_shift_avg'],
                            title='Päivittäinen kehitys',
                            labels={'value': 'Inc/työnt./h', 'variable': 'Vuoro'}
                        )
                        fig_daily.add_hline(y=5.1, line_dash="dash", line_color="red")
                        fig_daily.add_hline(y=4.6, line_dash="dash", line_color="blue")
                        st.plotly_chart(fig_daily, use_container_width=True)
                        
                        # Päivittäinen taulukko
                        st.subheader("Päivittäiset tulokset")
                        daily_display = daily_stats.copy()
                        daily_display['Päivätyöntekijät'] = daily_display.apply(
                            lambda x: f"{x['day_shift_avg']} {'✅' if x['day_target_met'] else '❌'}", axis=1
                        )
                        daily_display['Yötyöntekijät'] = daily_display.apply(
                            lambda x: f"{x['night_shift_avg']} {'✅' if x['night_target_met'] else '❌'}", axis=1
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
                
                with tab5:
                    st.subheader("💡 Optimointisuositukset")
                    
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
        
        except Exception as e:
            st.error(f"Virhe tiedoston käsittelyssä: {str(e)}")
            st.info("Tarkista että Excel-tiedosto sisältää sarakkeet 'Hour' ja 'Incidents handled by agent'.")
    
    else:
        # Ohjeet kun ei tiedostoa ladattu
        st.info("👆 Lataa Excel-tiedosto sivupalkista aloittaaksesi analyysin.")
        
        st.markdown("---")
        st.subheader("📋 Käyttöohjeet")
        st.markdown("""
        1. **Lataa Excel-tiedosto** sivupalkista
        2. Tiedoston tulee sisältää vähintään sarakkeet:
           - `Hour` (0-23)
           - `Incidents handled by agent` (määrä)
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

if __name__ == "__main__":
    main()