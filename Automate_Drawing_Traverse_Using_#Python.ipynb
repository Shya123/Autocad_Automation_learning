{
 "cells": [
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "<font face=\"Arial\" size=\"+2\" color=\"black\"><u><b>Automate Traverse Drawing Using </u></b><font face=\"Arial\" size=\"+3\" color=\"green\"><u><b>#</u></b><font face=\"Arial\" size=\"+2\" color=\"purple\"><u><b>Python</u></b><font face=\"Arial\" size=\"+3\" color=\"red\"><u><b>#</u></b><font face=\"Arial\" size=\"+2\" color=\"green\"><u><b>Pyautocad</u></b><font face=\"Arial\" size=\"+3\" color=\"orange\"><u><b>#</u></b><font face=\"Arial\" size=\"+2\" color=\"brown\"><u><b>Pandas</u></b>"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    "<font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>1.This code cell is the required package and library required to execute the task.</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 53,
   "metadata": {},
   "outputs": [],
   "source": [
    "from pyautocad import Autocad, APoint\n",
    "import pandas as pd"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    " <font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>2.This code cell links autocad modelspace for the drawing.</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": [
    "acad=Autocad()     ###create_if_not_exists=True\n",
    "print(acad.doc.Name)\n",
    "doc=acad.ActiveDocument\n",
    "ms=doc.ModelSpace"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    " <font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>3.This code cell will set layers details to be loaded on autocad modelspace.</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 55,
   "metadata": {},
   "outputs": [],
   "source": [
    "layers=['Pnt','ln','text']\n",
    "colors = [20,92,2]\n",
    "Lt= ['HIDDEN','ACAD_ISO06W100','PHANTOM']"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    " <font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>4.This code cell will load layers on autocad modelspace.</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 56,
   "metadata": {},
   "outputs": [],
   "source": [
    "for i in range(len(layers)):\n",
    "    Layer1=acad.ActiveDocument.Layers.Add(layers[i])\n",
    "    Layer1.color=colors[i]\n",
    "    acad.ActiveDocument.Linetypes.Load(Lt[i],\"acad.lin\")\n",
    "    Layer1.Linetype=Lt[i]\n",
    "    Layer1.Lineweight= 0.02"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    " <font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>5.Importing traverse points in Jupyter notebook .Here Pandas library is used.</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 57,
   "metadata": {},
   "outputs": [],
   "source": [
    "file=pd.read_excel('traverse_points.xlsx')  "
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    " <font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>6.Displaying the data from imported excel file.</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 58,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/html": [
       "<div>\n",
       "<style scoped>\n",
       "    .dataframe tbody tr th:only-of-type {\n",
       "        vertical-align: middle;\n",
       "    }\n",
       "\n",
       "    .dataframe tbody tr th {\n",
       "        vertical-align: top;\n",
       "    }\n",
       "\n",
       "    .dataframe thead th {\n",
       "        text-align: right;\n",
       "    }\n",
       "</style>\n",
       "<table border=\"1\" class=\"dataframe\">\n",
       "  <thead>\n",
       "    <tr style=\"text-align: right;\">\n",
       "      <th></th>\n",
       "      <th>S.N.</th>\n",
       "      <th>Easting</th>\n",
       "      <th>Northing</th>\n",
       "      <th>Elevation</th>\n",
       "      <th>Remarks</th>\n",
       "    </tr>\n",
       "  </thead>\n",
       "  <tbody>\n",
       "    <tr>\n",
       "      <th>0</th>\n",
       "      <td>1</td>\n",
       "      <td>618642.515</td>\n",
       "      <td>303483.980</td>\n",
       "      <td>1005.913</td>\n",
       "      <td>A</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>1</th>\n",
       "      <td>2</td>\n",
       "      <td>618585.367</td>\n",
       "      <td>303437.530</td>\n",
       "      <td>944.338</td>\n",
       "      <td>B</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>2</th>\n",
       "      <td>3</td>\n",
       "      <td>618711.664</td>\n",
       "      <td>303381.707</td>\n",
       "      <td>940.929</td>\n",
       "      <td>C</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>3</th>\n",
       "      <td>4</td>\n",
       "      <td>618799.685</td>\n",
       "      <td>303475.868</td>\n",
       "      <td>875.036</td>\n",
       "      <td>D</td>\n",
       "    </tr>\n",
       "    <tr>\n",
       "      <th>4</th>\n",
       "      <td>5</td>\n",
       "      <td>618741.575</td>\n",
       "      <td>303523.867</td>\n",
       "      <td>885.654</td>\n",
       "      <td>E</td>\n",
       "    </tr>\n",
       "  </tbody>\n",
       "</table>\n",
       "</div>"
      ],
      "text/plain": [
       "   S.N.     Easting    Northing  Elevation Remarks\n",
       "0     1  618642.515  303483.980   1005.913       A\n",
       "1     2  618585.367  303437.530    944.338       B\n",
       "2     3  618711.664  303381.707    940.929       C\n",
       "3     4  618799.685  303475.868    875.036       D\n",
       "4     5  618741.575  303523.867    885.654       E"
      ]
     },
     "execution_count": 58,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "file"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    " <font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>7.Finding number of data(or number of stations).</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 59,
   "metadata": {},
   "outputs": [
    {
     "data": {
      "text/plain": [
       "5"
      ]
     },
     "execution_count": 59,
     "metadata": {},
     "output_type": "execute_result"
    }
   ],
   "source": [
    "n=file['S.N.'].count()\n",
    "n"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    " <font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>8.This code cell will draw points on autocad modelspace and also list the points for connecting it with lines.</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 63,
   "metadata": {},
   "outputs": [],
   "source": [
    "points_list=[]\n",
    "for i in range(n):\n",
    "    points= APoint(file['Easting'][i],file['Northing'][i],file['Elevation'][i])\n",
    "    #print(\"Points no {} : x = {} , y = {} \".format(i+1,points[0],points[1]))\n",
    "    #print(points)\n",
    "    InsertionPnt=points\n",
    "    pointblock=acad.model.InsertBlock(InsertionPnt,'pointblck',7.5,7.5,7.5,0)\n",
    "    pointblock.layer = 'Pnt'\n",
    "    points_list.append(points)\n",
    "    \n",
    "    points_loc=APoint(points[0]-1,points[1]+3)\n",
    "    text_length=acad.model.AddText('%s'%(file['Remarks'][i]),points_loc,5)\n",
    "    text_length.layer='text'\n",
    "    text_with_coordinates=acad.model.AddText('   (%s mE, %s mN)'%(points[0],points[1]),points_loc,5)\n",
    "    text_with_coordinates.layer='text'"
   ]
  },
  {
   "cell_type": "markdown",
   "metadata": {},
   "source": [
    " <font face=\"Arial\" size=\"+1.7\" color=\"blue\"><u><b>9.This code cell will draw lines between points on autocad modelspace.</u><b>"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": 62,
   "metadata": {},
   "outputs": [],
   "source": [
    "for i in range(n):\n",
    "    if i< n-1:\n",
    "        line1=acad.model.AddLine(points_list[i],points_list[i+1])\n",
    "        line1.layer='ln'\n",
    "    else:\n",
    "        line1=acad.model.AddLine(points_list[-1],points_list[0])\n",
    "        line1.layer='ln'"
   ]
  },
  {
   "cell_type": "code",
   "execution_count": null,
   "metadata": {},
   "outputs": [],
   "source": []
  }
 ],
 "metadata": {
  "kernelspec": {
   "display_name": "Python 3",
   "language": "python",
   "name": "python3"
  },
  "language_info": {
   "codemirror_mode": {
    "name": "ipython",
    "version": 3
   },
   "file_extension": ".py",
   "mimetype": "text/x-python",
   "name": "python",
   "nbconvert_exporter": "python",
   "pygments_lexer": "ipython3",
   "version": "3.7.6"
  }
 },
 "nbformat": 4,
 "nbformat_minor": 4
}
