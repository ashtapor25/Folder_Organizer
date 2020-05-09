import setuptools

with open("README.md", "r") as fh:
    long_description = fh.read()

setuptools.setup(
    name="Folder_Organizer_Ashtapor25", # Replace with your own username
    version="0.0.1",
    author="Junsu Shin",
    author_email="junsushin4546@gmail.com",
    description="A Python package for organizing folders in a zip file into a specific hierarchy",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/ashtapor25/Folder_Organizer",
    packages=setuptools.find_packages(),
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved ::  GPL-3.0",
        "Operating System :: Windows 64bit",
    ],
    python_requires='>=3.6',
)
