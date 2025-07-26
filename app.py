import streamlit as st
import pandas as pd
import io
from datetime import datetime
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas

def create_label_pdf(product_name, label_size):
    """
    Create a PDF with labels based on the selected size
    
    Args:
        product_name (str): Name of the product
        label_size (str): Either "48x25mm" or "96x25mm"
    
    Returns:
        bytes: PDF content as bytes
    """
    buffer = io.BytesIO()
    
    # Convert mm to points (1 mm = 2.834645669 points)
    mm_to_pt = 2.834645669
    
    if label_size == "48x25mm":
        # Single label: 48mm x 25mm
        width = 48 * mm_to_pt
        height = 25 * mm_to_pt
        c = canvas.Canvas(buffer, pagesize=(width, height))
        draw_single_label(c, product_name, width, height)
    else:
        # Two labels side by side: 96mm x 25mm total (48mm x 25mm each)
        width = 96 * mm_to_pt
        height = 25 * mm_to_pt
        c = canvas.Canvas(buffer, pagesize=(width, height))
        
        # Draw two identical labels side by side
        label_width = 48 * mm_to_pt
        draw_single_label(c, product_name, label_width, height, x_offset=0)
        draw_single_label(c, product_name, label_width, height, x_offset=label_width)
    
    c.save()
    buffer.seek(0)
    return buffer.getvalue()

def draw_single_label(canvas_obj, product_name, width, height, x_offset=0.0):
    """
    Draw a single label on the canvas with improved sizing and spacing
    
    Args:
        canvas_obj: ReportLab canvas object
        product_name (str): Name of the product
        width (float): Width of the label in points
        height (float): Height of the label in points
        x_offset (float): Horizontal offset for positioning
    """
    # Get current date in DD/MM/YYYY format
    current_date = datetime.now().strftime("%d/%m/%Y")
    
    # Calculate dynamic font sizes based on label dimensions
    # For 48x25mm, use larger fonts to fill the space better
    base_width = 48 * 2.834645669  # 48mm in points
    scale_factor = width / base_width
    
    # Product name font size (large and bold) - increased size
    product_font_size = max(12, int(16 * scale_factor))
    # Date font size - increased size
    date_font_size = max(8, int(12 * scale_factor))
    
    # Calculate optimal spacing to fill the label space
    label_padding = height * 0.1  # 10% padding from top and bottom
    usable_height = height - (2 * label_padding)
    
    # Position product name in upper 60% of usable space
    product_name_y = label_padding + (usable_height * 0.7)
    # Position date in lower 30% of usable space  
    date_y = label_padding + (usable_height * 0.25)
    
    # Set font for product name (large and bold)
    canvas_obj.setFont("Helvetica-Bold", product_font_size)
    
    # Check if product name fits, if not, reduce font size
    text_width = canvas_obj.stringWidth(product_name, "Helvetica-Bold", product_font_size)
    available_width = width * 0.9  # Use 90% of width for text
    
    # Adjust font size if text is too wide
    while text_width > available_width and product_font_size > 8:
        product_font_size -= 1
        canvas_obj.setFont("Helvetica-Bold", product_font_size)
        text_width = canvas_obj.stringWidth(product_name, "Helvetica-Bold", product_font_size)
    
    # Draw product name (center-aligned)
    product_name_x = x_offset + (width - text_width) / 2
    canvas_obj.drawString(product_name_x, product_name_y, product_name)
    
    # Set font for date
    canvas_obj.setFont("Helvetica-Bold", date_font_size)
    
    # Draw date (center-aligned)
    date_text_width = canvas_obj.stringWidth(current_date, "Helvetica-Bold", date_font_size)
    date_x = x_offset + (width - date_text_width) / 2
    canvas_obj.drawString(date_x, date_y, current_date)
    
    # Draw border around the label with some padding
    border_padding = 2  # Small padding for border
    canvas_obj.rect(x_offset + border_padding, border_padding, 
                   width - (2 * border_padding), height - (2 * border_padding))

def load_google_sheet_data():
    """
    Load data from Google Sheets using public URL
    
    Returns:
        tuple: (list of unique product names, dataframe)
    """
    try:
        # Google Sheet URL (public access)
        sheet_url = "https://docs.google.com/spreadsheets/d/11dBw92P7Bg0oFyfqramGqdAlLTGhcb2ScjmR_1wtiTM/edit?usp=sharing"
        
        # Convert to CSV export URL
        csv_url = sheet_url.replace('/edit?usp=sharing', '/export?format=csv')
        
        # Read the data
        df = pd.read_csv(csv_url)
        
        if not df.empty:
            # Look for 'Name' column first (case-insensitive)
            name_column = None
            for col in df.columns:
                if col.lower().strip() in ['name', 'product name', 'product', 'item name', 'item']:
                    name_column = col
                    break
            
            if name_column:
                # Use the identified name column and remove duplicates/NaN
                product_names = df[name_column].dropna().unique().tolist()
                # Remove any empty strings and convert to strings
                product_names = [str(name).strip() for name in product_names if str(name).strip() and str(name).strip().lower() != 'nan']
                return product_names, df
            else:
                # Fallback to first column if no 'Name' column found
                product_names = df.iloc[:, 0].dropna().unique().tolist()
                product_names = [str(name).strip() for name in product_names if str(name).strip() and str(name).strip().lower() != 'nan']
                return product_names, df
        else:
            return [], df
    except Exception as e:
        st.error(f"Error loading Google Sheet: {str(e)}")
        st.error("Please check if the Google Sheet is publicly accessible and the URL is correct.")
        return [], pd.DataFrame()

def load_excel_file(uploaded_file):
    """
    Load Excel file and extract unique product names from 'Name' column
    (Backup method if Google Sheets fails)
    
    Args:
        uploaded_file: Streamlit uploaded file object
    
    Returns:
        tuple: (list of unique product names, list of column names, dataframe)
    """
    try:
        # Read Excel file
        df = pd.read_excel(uploaded_file)
        
        if not df.empty:
            # Look for 'Name' column first (case-insensitive)
            name_column = None
            for col in df.columns:
                if col.lower().strip() in ['name', 'product name', 'product', 'item name', 'item']:
                    name_column = col
                    break
            
            if name_column:
                # Use the identified name column and remove duplicates/NaN
                product_names = df[name_column].dropna().unique().tolist()
                # Remove any empty strings and convert to strings
                product_names = [str(name).strip() for name in product_names if str(name).strip() and str(name).strip().lower() != 'nan']
                return product_names, df.columns.tolist(), df
            else:
                # Fallback to first column if no 'Name' column found
                product_names = df.iloc[:, 0].dropna().unique().tolist()
                product_names = [str(name).strip() for name in product_names if str(name).strip() and str(name).strip().lower() != 'nan']
                return product_names, df.columns.tolist(), df
        else:
            return [], [], df
    except Exception as e:
        st.error(f"Error reading Excel file: {str(e)}")
        return [], [], pd.DataFrame()

def main():
    st.set_page_config(
        page_title="Product Label Generator",
        page_icon="🏷️",
        layout="centered"
    )
    st.title("🏷️ Product Label Generator")

    # Load data from Google Sheets
    with st.spinner("Loading product data..."):
        product_names, df = load_google_sheet_data()

    if product_names:
        # Minimal preview, no extra info
        with st.expander(f"Show all products ({len(product_names)})"):
            cols = st.columns(3)
            for i, name in enumerate(product_names):
                col_idx = i % 3
                cols[col_idx].write(f"{i+1}. {name}")

        selected_product = st.selectbox(
            "Select product:",
            options=product_names,
            index=0
        )

        label_size = st.radio(
            "Label size:",
            options=["48x25mm", "96x25mm"]
        )

        # Preview section
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"<div style='text-align: center; font-size: 20px; font-weight: bold; padding: 10px; border: 2px solid #ddd; border-radius: 5px; margin: 10px 0;'>{selected_product}</div>", unsafe_allow_html=True)
        with col2:
            current_date = datetime.now().strftime("%d/%m/%Y")
            st.markdown(f"<div style='text-align: center; font-size: 16px; font-weight: bold; padding: 10px; border: 2px solid #ddd; border-radius: 5px; margin: 10px 0;'>{current_date}</div>", unsafe_allow_html=True)

        if st.button("Generate Label PDF", type="primary"):
            try:
                pdf_bytes = create_label_pdf(selected_product, label_size)
                safe_product_name = str(selected_product).replace(' ', '_').replace('/', '_').replace('\\', '_')
                filename = f"{safe_product_name}_{label_size}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                st.download_button(
                    label="Download PDF Label",
                    data=pdf_bytes,
                    file_name=filename,
                    mime="application/pdf",
                    type="secondary"
                )
            except Exception as e:
                st.error(f"Error generating PDF: {str(e)}")

        # Minimal note about data source in sidebar
        st.sidebar.markdown("<small>Data source: <a href='https://docs.google.com/spreadsheets/d/11dBw92P7Bg0oFyfqramGqdAlLTGhcb2ScjmR_1wtiTM' target='_blank'>Google Sheet</a></small>", unsafe_allow_html=True)
    
    else:
        # Fallback to Excel upload if Google Sheets fails
        st.warning("⚠️ Could not load data from Google Sheets. Using Excel upload as backup.")
        
        # File upload section
        st.header("📁 Upload Excel File (Backup Method)")
        uploaded_file = st.file_uploader(
            "Choose an Excel file containing product names",
            type=["xlsx", "xls"],
            help="The first column should contain product names"
        )
        
        if uploaded_file is not None:
            # Load product names from Excel
            product_names, columns, df = load_excel_file(uploaded_file)
            
            if product_names:
                st.success(f"✅ Found {len(product_names)} unique products in the Excel file")
                
                # Show which column was used
                name_column = None
                for col in df.columns:
                    if col.lower().strip() in ['name', 'product name', 'product', 'item name', 'item']:
                        name_column = col
                        break
                
                if name_column:
                    st.info(f"📊 Using '{name_column}' column for product names (duplicates removed)")
                else:
                    st.info(f"📊 Using first column ('{df.columns[0]}') for product names (duplicates removed)")
                
                # Display first few product names as preview
                st.subheader("📋 Product Preview")
                with st.expander(f"Click to see all {len(product_names)} unique products"):
                    # Display in columns for better layout
                    cols = st.columns(3)
                    for i, name in enumerate(product_names):
                        col_idx = i % 3
                        cols[col_idx].write(f"{i+1}. {name}")
                
                # Product selection section
                st.header("🎯 Select Product")
                selected_product = st.selectbox(
                    "Choose a product to generate label for:",
                    options=product_names,
                    index=0,
                    help="This dropdown shows only unique product names from your Excel file"
                )
                
                # Rest of the interface (same as Google Sheets version)
                # Label size selection
                st.header("📏 Select Label Size")
                label_size = st.radio(
                    "Choose label dimensions:",
                    options=["48x25mm", "96x25mm"],
                    help="48x25mm: Single label | 96x25mm: Two identical labels side by side"
                )
                
                # Preview section
                st.header("👀 Label Preview")
                col1, col2 = st.columns(2)
                
                with col1:
                    st.write("**Product Name:**")
                    st.markdown(f"<div style='text-align: center; font-size: 20px; font-weight: bold; padding: 10px; border: 2px solid #ddd; border-radius: 5px; margin: 10px 0;'>{selected_product}</div>", 
                               unsafe_allow_html=True)
                
                with col2:
                    st.write("**Date:**")
                    current_date = datetime.now().strftime("%d/%m/%Y")
                    st.markdown(f"<div style='text-align: center; font-size: 16px; font-weight: bold; padding: 10px; border: 2px solid #ddd; border-radius: 5px; margin: 10px 0;'>{current_date}</div>", 
                               unsafe_allow_html=True)
                
                # Label size info with visual representation
                if label_size == "48x25mm":
                    st.info("📄 Single label (48mm × 25mm) - Larger fonts for better visibility")
                else:
                    st.info("📄📄 Two identical labels side by side (96mm × 25mm total) - Perfect for efficiency")
                
                # Generate and download button
                st.header("⬇️ Generate & Download")
                if st.button("🎯 Generate Label PDF", type="primary"):
                    try:
                        # Generate PDF
                        pdf_bytes = create_label_pdf(selected_product, label_size)
                        
                        # Create download button
                        safe_product_name = str(selected_product).replace(' ', '_').replace('/', '_').replace('\\', '_')
                        filename = f"{safe_product_name}_{label_size}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
                        
                        st.download_button(
                            label="📥 Download PDF Label",
                            data=pdf_bytes,
                            file_name=filename,
                            mime="application/pdf",
                            type="secondary"
                        )
                        
                        st.success("✅ Label generated successfully! Click the download button above.")
                        
                    except Exception as e:
                        st.error(f"❌ Error generating PDF: {str(e)}")
            
            else:
                st.warning("⚠️ No product names found in the uploaded file. Please check that the first column contains product names.")
        
        else:
            # Instructions when no file is uploaded and Google Sheets failed
            st.info("👆 Please upload an Excel file to get started")
            
            st.header("📖 Instructions")
            st.markdown("""
            **Primary Method (Automatic):**
            - The app automatically loads data from Google Sheets
            - Data is always up-to-date with your online spreadsheet
            
            **Backup Method (Manual Upload):**
            1. **Prepare your Excel file**: Ensure product names are in the first column
            2. **Upload the file**: Use the file uploader above
            3. **Select a product**: Choose from the dropdown list
            4. **Choose label size**: 
               - 48x25mm: Single label
               - 96x25mm: Two identical labels side by side
            5. **Generate & Download**: Click the button to create your PDF label
            
            **Label Format:**
            - Product name: Large, bold, center-aligned
            - Date: Auto-filled with current date (DD/MM/YYYY format)
            - Professional PDF output with exact dimensions
            """)
    
    # Add refresh button for Google Sheets data
    st.sidebar.header("Data Management")
    if st.sidebar.button("Refresh Google Sheets Data"):
        st.rerun()

if __name__ == "__main__":
    main()