# -*- coding: utf-8 -*-

"""
Zaim の記録データを 文字コード=Shift-JIS 形式でダウンロードする
日付,方法,カテゴリ,カテゴリの内訳,支払元,入金先,品目,メモ,お店,通貨,収入,支出,振替,残高調整,通貨変換前の金額,集計の設定
"""

import os
import csv
import sys
import glob
import codecs
import traceback


if __name__ == '__main__':


    def get_html_line_tmp():

        html_line_tmp = u'''\
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Zaim</title>
<head>
<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
<script type="text/javascript">google.charts.load("current", {{packages:["timeline", "corechart"]}});</script>

<script type="text/javascript">
google.charts.setOnLoadCallback(drawChart);
function drawChart() {{

    var data = google.visualization.arrayToDataTable([
{datasP}
    ]);

    var options = {{
        width: 1800,
        height: 600,
        legend: {{ position: 'none' }},
        bar: {{ groupWidth: '50%' }},
        isStacked: true,
    }};
    var chart = new google.visualization.ColumnChart(document.getElementById('ZaimP'));
    chart.draw(data, options);
    }}
</script>

<script type="text/javascript">
google.charts.setOnLoadCallback(drawChart);
function drawChart() {{

    var data = google.visualization.arrayToDataTable([
{datasC}
    ]);

    var options = {{
        width: 1800,
        height: 600,
        legend: {{ position: 'none' }},
        bar: {{ groupWidth: '50%' }},
        isStacked: true,
    }};
    var chart = new google.visualization.ColumnChart(document.getElementById('ZaimC'));
    chart.draw(data, options);
    }}
</script>

</head>
<body>
    <div id="ZaimP"></div>
    <div id="ZaimC"></div>
</body>
</html>
'''
        return html_line_tmp


    def get_html_timeline_data_line_tmp():

        timeline_data_line_tmp = u'''\

'''
        return timeline_data_line_tmp


    def get_category_label( line ):

        category_main = u'{0}'.format( line[2].decode( 'shift-jis', 'ignore' ) ) # カテゴリ
        category_sub = u'{0}'.format( line[3].decode( 'shift-jis', 'ignore' ) ) # カテゴリの内訳

        return u'{0}:{1}'.format( category_main, category_sub )


    try:

        cur_dir_path =  os.path.abspath( r'.' )

        csv_file_path_list = glob.glob( os.path.join( cur_dir_path, r'*.csv' ) )

        for csv_file_path in csv_file_path_list:

            lines = []

            # open .csv
            with open( csv_file_path ) as f:
                for line in csv.reader( f ):
                    lines.append( line )

            # column label を取得
            category_label_list = []
            payment_label_list = []

            for line in lines[1:]: # 0 = columun 列 除外
                category_label_list.append( get_category_label( line ) )
                payment_label_list.append( u'{0}'.format( line[4].decode( 'shift-jis', 'ignore' ) ) )

            category_label_list = [ item for item in set( category_label_list ) ]
            category_label_list = sorted( category_label_list )

            payment_label_list = [ item for item in set( payment_label_list ) ]
            payment_label_list = sorted( payment_label_list )


            payment_src_items_dict = {} # 口座別
            category_items_dict = {} # カテゴリ別

            for line in lines[1:]: # 0 = columun 列 除外

                # get line data
                method = line[1] # 方法

                payment_src = u'{0}'.format( line[4].decode( 'shift-jis', 'ignore' ) ) # 支払元
                category_label = get_category_label( line ) # カテゴリ

                #income_price = line[10] # 収入
                payment_price = int( line[11] ) # 支出

                # 口座別
                if not payment_src_items_dict.has_key( payment_src ):
                    payment_src_items_dict[ payment_src ] = [ 0.0 ] * len( category_label_list )
                payment_src_items_dict[ payment_src ][ category_label_list.index( category_label ) ] += payment_price

                # カテゴリ別
                if not category_items_dict.has_key( category_label ):
                    category_items_dict[ category_label ] = [ 0.0 ] * len( payment_label_list )
                category_items_dict[ category_label ][ payment_label_list.index( payment_src ) ] += payment_price

            # 口座別
            category_label_list[0:0] = [ u'カテゴリ' ] # column の先頭に追加

            payment_data_row_list = []
            payment_data_row_list.append(
                u'[{0}]'.format( ','.join( u"'{0}'".format( item ) for item in category_label_list ) )
            )

            payment_src_key_list = payment_src_items_dict.keys()
            payment_src_key_list = sorted( payment_src_key_list )

            for payment_src_key in payment_src_key_list:

                payment_list = payment_src_items_dict[ payment_src_key ]
                if sum( payment_list ) == 0.0: # No payment
                    continue

                row_item_list = [ u'{0}'.format( item ) for item in payment_list ]
                row_item_list[0:0] = [ u"'{0}'".format( payment_src_key ) ] # column の先頭に追加

                payment_data_row_list.append( u'[{0}]'.format( ','.join( row_item_list ) ) )

            # カテゴリ
            payment_label_list[0:0] = [ u'口座' ] # column の先頭に追加

            category_data_row_list = []
            category_data_row_list.append(
                u'[{0}]'.format( ','.join( u"'{0}'".format( item ) for item in payment_label_list ) )
            )

            category_src_key_list = category_items_dict.keys()
            category_src_key_list = sorted( category_src_key_list )

            for category_src_key in category_src_key_list:

                payment_list = category_items_dict[ category_src_key ]
                if sum( payment_list ) == 0.0: # No payment
                    continue

                row_item_list = [ u'{0}'.format( item ) for item in payment_list ]
                row_item_list[0:0] = [ u"'{0}'".format( category_src_key ) ] # column の先頭に追加

                category_data_row_list.append( u'[{0}]'.format( ','.join( row_item_list ) ) )

            # html
            html_line = get_html_line_tmp().format(
                **{
                    'datasP': ',\n'.join( payment_data_row_list ),
                    'datasC': ',\n'.join( category_data_row_list )
                }
            )

            basename = os.path.basename( csv_file_path )
            basename, ext = os.path.splitext( basename )

            html_file_path = os.path.abspath( '{0}.html'.format( basename ) )
            with codecs.open( html_file_path, 'w', 'utf-8' ) as f:
                f.write( html_line )

            os.startfile( html_file_path )

    except:
        print traceback.format_exc()
        raw_input( '--- end ---' )
