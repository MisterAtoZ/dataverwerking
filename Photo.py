from PIL import Image, ImageDraw, ImageFont

class Photo():
    def makeTitle(title):
        img = Image.new('RGB', (6000, 4000), color = (73,109,137))

        fnt = ImageFont.truetype('C:/Windows/Fonts/Arial.ttf', 1000)
        d = ImageDraw.Draw(img)
        d.text((2000,1500), title, font=fnt, fill=(255,255,255))

        img.save(title + '.jpg')
        Photo.resize(title + '.jpg', title)
        return title + '.jpg'


    def resize(img, title):
        img = Image.open(img)
        basewidth = 300
        wpercent = (basewidth / float(img.size[0]))
        hsize = int((float(img.size[1]) * float(wpercent)))
        img = img.resize((basewidth, hsize), Image.ANTIALIAS)
        img.save(title + '.jpg')

    """ 
    merge_image takes three parameters first two parameters specify 
    the two images to be merged and third parameter i.e. vertically
    is a boolean type which if True merges images vertically
    and finally saves and returns the file_name
    """
    def merge_image(img1, img2, vertically, name):
        images = list(map(Image.open, [img1, img2]))
        widths, heights = zip(*(i.size for i in images))
        if vertically:
            max_width = max(widths)
            total_height = sum(heights)
            new_im = Image.new('RGB', (max_width, total_height))

            y_offset = 0
            for im in images:
                new_im.paste(im, (0, y_offset))
                y_offset += im.size[1]
        else:
            total_width = sum(widths)
            max_height = max(heights)
            new_im = Image.new('RGB', (total_width, max_height))

            x_offset = 0
            for im in images:
                new_im.paste(im, (x_offset, 0))
                x_offset += im.size[0]

        new_im.save(name + '.jpg')
        return name + '.jpg'

"""
title = Photo.makeTitle('0 h')
photo1 = Photo.merge_image(title, '0_JW1_F.jpg', True, 'c1')
photo2 = Photo.merge_image(photo1, '0_JW1_B.jpg', True, 'c1')
photo3 = Photo.merge_image(photo2, '0_JW2_F.jpg', True, 'c1')
photo4 = Photo.merge_image(photo3, '0_JW2_B.jpg', True, 'c1')


title = Photo.makeTitle('JW')
JW1_F = Photo.makeTitle('JW1_F')
JW1_B = Photo.makeTitle('JW1_B')
JW2_F = Photo.makeTitle('JW2_F')
JW2_B = Photo.makeTitle('JW2_B')

a1 = Photo.merge_image(title, JW1_F, True, 'c0')
a2 = Photo.merge_image(a1, JW1_B, True, 'c0')
a3 = Photo.merge_image(a2, JW2_F, True, 'c0')
a4 = Photo.merge_image(a3, JW2_B, True, 'c0')

tot = Photo.merge_image(a4, photo4, False, 'resultaat')
"""