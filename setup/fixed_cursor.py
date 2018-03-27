'''
Fixed Cursor
======
Ths is a changed version of the module Cursor.
This module still shows the mouse even if its outsite the window.
It helps if the mouse is lost.


Shows a cursor following mouse motion events, useful on systems with no
visible native mouse cursor.

Configuration
-------------

:Parameters:
    `texture`: str, defaults to
        'data/images/cursor.png' Image used to represent the cursor if
        displayed
    `size`: tuple, defaults to (40, 40)
        Apparent size of the mouse cursor, if displayed, (None,None) value
        will keep its real size.
    `offset`: tuple, defaults to (None, None)
        Offset of the texture image. The default value will align the
        top-left corner of the image to the mouse pos.

Example
-------

In your configuration (`~/.kivy/config.ini`), you can add something like
this::

    [modules]
    cursor = texture=mypointer.png,size=20x20,offset=20x20

.. versionadded:: 1.10.0
'''

__all__ = ('start', 'stop')

from kivy.core.image import Image
from kivy.graphics import Color, Rectangle
from kivy import kivy_data_dir
from kivy.compat import string_types
from os.path import join
from functools import partial

def _mouse_move(texture, size, offset, win, pos, *args):
    if hasattr(win, '_cursor'):
        c = win._cursor
    else:
        with win.canvas.after:
            Color(1, 1, 1, 1, mode='rgba')
            win._cursor = c = Rectangle(texture=texture, size=size)
    
    pos = [pos[0] + offset[0], pos[1] - size[1] + offset[1]]
    x_max, y_max = win.size
    x_max += - size[0] + offset[0]
    y_max += - size[1] + offset[1]
    if pos[0] < 0:
        pos[0] = 0
    elif pos[0] > x_max:
        pos[0] = x_max
    if pos[1] < 0:
        pos[1] = 0
    elif pos[1] > y_max:
        pos[1] = y_max
    c.pos = pos[0],pos[1]
    

def start(win, ctx):
    cursor_texture = Image(
        ctx.config.get('texture', join(kivy_data_dir, 'images', 'cursor.png'))
    ).texture
    cursor_size = ctx.config.get('size')
    if isinstance(cursor_size, string_types):
        cursor_size = [int(x) for x in cursor_size.split('x')]
    elif not cursor_size:
        cursor_size = cursor_texture.size

    cursor_offset = ctx.config.get('offset', (0, 0))
    if isinstance(cursor_offset, string_types):
        cursor_offset = [int(x) for x in cursor_offset.split('x')]

    win.bind(
        mouse_pos=partial(
            _mouse_move, cursor_texture, cursor_size, cursor_offset))


def stop(win, ctx):
    win.unbind(mouse_pos=_mouse_move)
