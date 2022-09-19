from setuptools import setup

setup(
    name='pymorningstar', 
    version='0.10', 
    description = "Automating the download of information from Morningstar's Excel Add-In API",
    author = 'Pablo Vilas',
    author_email='pablo.vilas.naval@gmail.com',
    packages = ['pymorningstar'],
    install_requires=['xlwings','pandas','pyautogui', 'opencv-python'],
    url='https://github.com/VilasPablo/2021-Morningstar-Excel',
    classifiers=[ 
        'Intended Audience :: Science/Research',
        'Intended Audience :: Financial and Insurance Industry',
        'Programming Language :: Python :: 3',  ],
    include_package_data=True,
    zip_safe = False)


# MANIFEST.in