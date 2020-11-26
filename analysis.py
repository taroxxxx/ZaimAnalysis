# -*- coding: utf-8 -*-

"""
Zaim の記録データを 文字コード=Shift-JIS 形式でダウンロードする
日付,方法,カテゴリ,カテゴリの内訳,支払元,入金先,品目,メモ,お店,通貨,収入,支出,振替,残高調整,通貨変換前の金額,集計の設定
"""

import os
import csv
import sys
import copy
import glob
import codecs
import traceback


if __name__ == '__main__':


    def to_utf( text ):
        return u'{0}'.format( text.decode( 'shift-jis', 'ignore' ) )


    def get_html_line_tmp():

        return u'''\
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Zaim</title>
<head>
<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
<script type="text/javascript">google.charts.load("current", {{packages:["timeline", "corechart"]}});</script>
{script_lines}
</head>
<body>
{div_lines}
</body>
</html>'''


    def get_html_chart_script_line_tmp():

        return u'''\
<script type="text/javascript">
google.charts.setOnLoadCallback(drawChart);
function drawChart() {{
    var data = google.visualization.arrayToDataTable([
{datas}
    ]);
    var options = {{
        title: '{title}',
        width: {width},
        height: {height},
        legend: {{ position: 'none' }},
        bar: {{ groupWidth: '75%' }},
        isStacked: true,
    }};
    var chart = new google.visualization.ColumnChart(document.getElementById('{id}'));
    chart.draw(data, options);
    }}
</script>'''


    def get_html_div_line_tmp():

        return u'''\
    <div id="{id}"></div>'''


    def get_category_label( line ):

        category_main = to_utf( line[2] ) # カテゴリ
        category_sub = to_utf( line[3] ) # 内訳

        return u'{0}:{1}'.format( category_main, category_sub )


    def set_sorted_list( tmp_list ):

        tmp_list = [ item for item in set( tmp_list ) ]
        tmp_list = sorted( tmp_list )

        return tmp_list


    try:
        # get .csv
        cur_dir_path =  os.path.abspath( r'.' )
        csv_file_path_list = glob.glob( os.path.join( cur_dir_path, r'*.csv' ) )

        for csv_file_path in csv_file_path_list:

            # read .csv
            lines = []
            with open( csv_file_path ) as f:
                for line in csv.reader( f ):
                    lines.append( line )

            # column label を取得
            payment_label_list = []
            category_label_list = []
            category_main_label_list = []

            for line in lines[1:]: # 0 = columun 列 除外
                payment_label_list.append( to_utf( line[4] ) ) # 支払元
                category_label_list.append( get_category_label( line ) ) # カテゴリ:内訳
                category_main_label_list.append( to_utf( line[2] ) ) # カテゴリ

            payment_label_list = set_sorted_list( payment_label_list )
            category_label_list = set_sorted_list( category_label_list )
            category_main_label_list = set_sorted_list( category_main_label_list )

            # data の取得
            category_items_dict = {} # カテゴリ:内訳
            payment_src_items_dict = {} # 口座

            category_main_items_items_dict = {} # カテゴリ

            for line in lines[1:]: # 0 = columun 列 除外

                # get line data
                method = line[1] # 方法
                category_main = to_utf( line[2] ) # カテゴリ
                category_sub = to_utf( line[3] ) # 内訳
                payment_src = to_utf( line[4] ) # 支払元
                income_price = int( line[10] ) # 収入
                payment_price = int( line[11] ) # 支出

                category_label = get_category_label( line ) # カテゴリ:内訳

                # カテゴリ:内訳
                if not category_items_dict.has_key( category_label ):
                    init_list = [ 0.0 ] * len( payment_label_list )
                    category_items_dict[ category_label ] = init_list

                index = payment_label_list.index( payment_src )
                category_items_dict[ category_label ][ index ] += payment_price

                # 口座
                if not payment_src_items_dict.has_key( payment_src ):
                    init_list = [ 0.0 ] * len( category_label_list )
                    payment_src_items_dict[ payment_src ] = init_list

                index = category_label_list.index( category_label )
                payment_src_items_dict[ payment_src ][ index ] += payment_price

                # カテゴリ
                if not category_main_items_items_dict.has_key( category_main ):
                    category_main_items_items_dict[ category_main ] = {}

                if not category_main_items_items_dict[ category_main ].has_key( category_sub ):
                    init_list = [ 0.0 ] * len( payment_label_list )
                    category_main_items_items_dict[ category_main ][ category_sub ] = init_list

                index = payment_label_list.index( payment_src )
                category_main_items_items_dict[ category_main ][ category_sub ][ index ] += payment_price

            # ColumnChart
            script_line_list = []
            div_line_list = []

            # カテゴリ ごとに分ける
            category_main_data_list = []
            for category_main_key in category_main_items_items_dict:
                category_main_data_list.append(
                    [
                        category_main_key,
                        payment_label_list,
                        category_main_items_items_dict[ category_main_key ],
                        450
                    ]
                )


            for id, item in enumerate(
                [
                    [ u'全カテゴリ', payment_label_list, category_items_dict, 900 ],
                    #[ u'口座', category_label_list, payment_src_items_dict ],
                ] + category_main_data_list
            ):

                input_first_label = item[0]
                input_label_list = item[1]
                input_items_dict = item[2]
                height = item[3]

                label_list = [ input_first_label ] # column の先頭に追加
                label_list.extend( input_label_list )

                data_row_list = []
                data_row_list.append(
                    u'        [{0}]'.format( ','.join( u"'{0}'".format( item ) for item in label_list ) )
                )

                items_dict = copy.copy( input_items_dict )

                src_key_list = items_dict.keys()
                src_key_list = sorted( src_key_list )

                total_payment = 0.0

                for src_key in src_key_list:

                    payment_list = items_dict[ src_key ]

                    if sum( payment_list ) == 0.0: # No payment
                        continue

                    row_item_list = [ u'{0}'.format( item ) for item in payment_list ]
                    row_item_list[0:0] = [ u"'{0}'".format( src_key ) ] # column の先頭に追加

                    data_row_list.append( u'        [{0}]'.format( ','.join( row_item_list ) ) )

                    total_payment += sum( payment_list )

                if len( data_row_list ) <= 1: # column row しかない場合
                    continue

                script_line_list.append( get_html_chart_script_line_tmp().format(
                    **{
                        'title': u'{0}: ￥{1}'.format( input_first_label, int(total_payment) ),
                        'datas': u',\n'.join( data_row_list ),
                        'id': u'data{0:02d}'.format( id ),
                        'width': 1800,
                        'height': height,
                    }
                ) )

                div_line_list.append( get_html_div_line_tmp().format(
                    **{
                        'id': u'data{0:02d}'.format( id ),
                    }
                ) )

            # html
            html_line = get_html_line_tmp().format(
                **{
                    'script_lines': '\n'.join( script_line_list ),
                    'div_lines': '\n'.join( div_line_list ),
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
