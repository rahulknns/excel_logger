from setuptools import setup
  
setup(
    name='excel_logger',
    version='1.0',
    description='provides class to log excel data into multiple sheets',
    author='Rahulkannan S',
    author_email='rahulknns@gmail.com',
    packages=['logger'],
    package_dir={'logger': 'src/logger'},

    install_requires=[
        'openpyxl'
    ],
)
