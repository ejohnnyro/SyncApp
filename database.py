from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, ForeignKey, Boolean
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, relationship
from datetime import datetime

Base = declarative_base()

class Product(Base):
    __tablename__ = 'products'
    
    id = Column(Integer, primary_key=True)
    woo_id = Column(Integer, unique=True)  # WooCommerce product ID
    name = Column(String)
    sku = Column(String, nullable=True)  # Removed unique constraint to handle duplicate SKUs
    regular_price = Column(Float, nullable=True)
    sale_price = Column(Float, nullable=True)
    stock_quantity = Column(Integer, nullable=True)
    categories = Column(String, nullable=True)  # Store category names as comma-separated string
    last_synced = Column(DateTime, default=datetime.now)  # Changed to use local time
    vendor_stocks = relationship('VendorStock', back_populates='product')
    
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.woo_id = kwargs.get('woo_id')
        self.name = kwargs.get('name')
        self.sku = kwargs.get('sku')
        self.regular_price = kwargs.get('regular_price')
        self.sale_price = kwargs.get('sale_price')
        self.stock_quantity = kwargs.get('stock_quantity')
        self.categories = kwargs.get('categories')
        self.last_synced = kwargs.get('last_synced', datetime.now())

class Vendor(Base):
    __tablename__ = 'vendors'
    
    id = Column(Integer, primary_key=True)
    name = Column(String, unique=True)
    api_url = Column(String)
    api_key = Column(String)
    api_secret = Column(String)
    is_active = Column(Boolean, default=True)
    last_sync = Column(DateTime)
    stocks = relationship('VendorStock', back_populates='vendor')

class VendorStock(Base):
    __tablename__ = 'vendor_stocks'
    
    id = Column(Integer, primary_key=True)
    product_id = Column(Integer, ForeignKey('products.id'))
    vendor_id = Column(Integer, ForeignKey('vendors.id'))
    stock_quantity = Column(Integer)
    price = Column(Float)
    last_updated = Column(DateTime, default=datetime.now)
    product = relationship('Product', back_populates='vendor_stocks')
    vendor = relationship('Vendor', back_populates='stocks')

class DatabaseManager:
    def __init__(self, db_path='sqlite:///products.db'):
        self.engine = create_engine(db_path)
        self.Session = sessionmaker(bind=self.engine)
        Base.metadata.create_all(self.engine)
    
    def get_session(self):
        return self.Session()
    
    def add_or_update_product(self, product_data):
        session = self.get_session()
        try:
            product = session.query(Product).filter_by(woo_id=product_data['id']).first()
            
            if not product:
                product = Product()
            
            # Update product attributes
            product.woo_id = product_data['id']
            product.name = product_data['name']
            product.sku = product_data.get('sku', '')
            product.regular_price = float(product_data.get('regular_price', 0)) if product_data.get('regular_price') else None
            product.sale_price = float(product_data.get('sale_price', 0)) if product_data.get('sale_price') else None
            product.stock_quantity = int(product_data.get('stock_quantity', 0)) if product_data.get('stock_quantity') else None
            # Convert categories list to comma-separated string
            categories = product_data.get('categories', [])
            if categories and isinstance(categories, list):
                product.categories = ', '.join(cat['name'] for cat in categories)
            else:
                product.categories = None
                
            product.last_synced = datetime.now()  # Changed to use local time
            
            if product.id is None:
                session.add(product)
            session.commit()
            return product.id
        except Exception as e:
            session.rollback()
            raise e
        finally:
            session.close()
    
    def get_product(self, product_id):
        session = self.get_session()
        try:
            return session.query(Product).filter_by(id=product_id).first()
        finally:
            session.close()
    
    def search_products(self, search_term=None, limit=100, offset=0):
        session = self.get_session()
        try:
            query = session.query(Product)
            if search_term:
                search_term = f'%{search_term}%'
                query = query.filter(
                    (Product.name.ilike(search_term)) |
                    (Product.sku.ilike(search_term))
                )
            return query.limit(limit).offset(offset).all()
        finally:
            session.close()
    
    def add_vendor(self, name, api_url, api_key, api_secret):
        session = self.get_session()
        try:
            vendor = Vendor(
                name=name,
                api_url=api_url,
                api_key=api_key,
                api_secret=api_secret
            )
            session.add(vendor)
            session.commit()
            return vendor.id
        except Exception as e:
            session.rollback()
            raise e
        finally:
            session.close()
    
    def update_vendor_stock(self, product_id, vendor_id, quantity, price):
        session = self.get_session()
        try:
            stock = session.query(VendorStock).filter_by(
                product_id=product_id,
                vendor_id=vendor_id
            ).first()
            
            if not stock:
                stock = VendorStock(
                    product_id=product_id,
                    vendor_id=vendor_id
                )
                session.add(stock)
            
            stock.stock_quantity = quantity
            stock.price = price
            stock.last_updated = datetime.now()
            session.commit()
            return stock.id
        except Exception as e:
            session.rollback()
            raise e
        finally:
            session.close()
    
    def get_total_products(self):
        session = self.get_session()
        try:
            return session.query(Product).count()
        finally:
            session.close()
    
    def update_product_field(self, product_id, field_name, value):
        session = self.get_session()
        try:
            product = session.query(Product).filter_by(woo_id=product_id).first()
            if product:
                setattr(product, field_name, value)
                product.last_synced = datetime.now()
                session.commit()
                return True
            return False
        except Exception as e:
            session.rollback()
            raise e
        finally:
            session.close()