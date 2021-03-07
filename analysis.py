# -*- coding: utf-8 -*-

"""
Zaim の記録データを 文字コード=Shift-JIS 形式でダウンロードする
日付,方法,カテゴリ,カテゴリの内訳,支払元,入金先,品目,メモ,お店,通貨,収入,支出,振替,残高調整,通貨変換前の金額,集計の設定
"""

import os
import re
import csv
import sys
import copy
import glob
import time
import datetime
import codecs
import traceback

from operator import itemgetter


if __name__ == '__main__':


    def to_utf( text ):
        return u'{0}'.format( text.decode( 'shift-jis', 'ignore' ) )


    def get_html_line_tmp():

        return u'''\
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>Zaim: {date}</title>
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


    def get_transfer_label( line ):

        payment_src = to_utf( line[4] ) # 支払元
        income_dst = to_utf( line[5] ) # 入金先

        return u'{0}>{1}'.format( payment_src, income_dst )


    def set_sorted_list( tmp_list ):

        tmp_list = [ item for item in set( tmp_list ) ]
        tmp_list = sorted( tmp_list )

        return tmp_list


    def date_str_to_datetime( date_str ):

        date_str_fmt = re.compile( r'(?P<y>[\d]+)-(?P<m>[\d]+)-(?P<d>[\d]+)' )
        date_str_search = date_str_fmt.search( date_str )

        if date_str_search != None:

            y = int( date_str_search.group( 'y' ) )
            m = int( date_str_search.group( 'm' ) )
            d = int( date_str_search.group( 'd' ) )

            return datetime.datetime( y,m,d )

        return None


    def datetime_to_time( src_d ):

        """
        datetime@src_d
        """
        return time.mktime( src_d.timetuple() )


    def src_time_to_str( src_time, tmp=u'{year}/{month}/{day}' ):

        """
        # 入力時間をテキストで返す

        i@src_time : time
        s@tmp : テキストテンプレート
        """
        d = datetime.datetime.fromtimestamp( src_time )

        fmt_dict = {
            'year' : d.year,
            'month' : d.month,
            'day' : d.day,
            'weekday_utf' : [u'月',u'火',u'水',u'木',u'金',u'土',u'日'][d.weekday()],
        }

        return tmp.format( **fmt_dict ) if src_time else '-'*5

    try:

        def ___main___():
            pass

        # get .csv
        cur_dir_path =  os.path.abspath( r'.' )
        csv_file_path_list = glob.glob( os.path.join( cur_dir_path, r'*.csv' ) )

        for csv_file_path in csv_file_path_list:

            # read .csv
            lines = []
            with open( csv_file_path ) as f:
                for line in csv.reader( f ):
                    lines.append( line )

            """ Columun を取得 """
            payment_label_list = []
            category_label_list = []
            category_main_label_list = []
            income_label_list = []
            transfer_label_list = []

            for line in lines[1:]: # 0 = columun 列 除外

                payment_label_list.append( to_utf( line[4] ) ) # 支払元
                category_main_label_list.append( to_utf( line[2] ) ) # カテゴリ
                category_label_list.append( get_category_label( line ) ) # カテゴリ:内訳
                income_label_list.append( to_utf( line[5] ) ) # 入金先
                transfer_label_list.append( get_transfer_label( line ) ) # 振替: 支払元>入金先

            payment_label_list = set_sorted_list( payment_label_list )
            income_label_list = set_sorted_list( income_label_list )
            transfer_label_list = set_sorted_list( transfer_label_list )
            category_label_list = set_sorted_list( category_label_list )
            category_main_label_list = set_sorted_list( category_main_label_list )

            """ line ごとに集計 """
            category_items_dict = {} # 支出: カテゴリ:内訳:
            payment_src_items_dict = {} # 支出: 口座
            category_main_items_items_dict = {} # 支出: カテゴリ
            income_dst_items_dict = {} # 収入: 口座:
            transfer_items_dict = {} # 振替

            specified_category_main_list = [ u'食費' ]
            specified_category_sub_list = [ u'外食(昼)', u'間食', u'飲料' ]
            specified_payment_label = u'和'
            specified_payment_label_list = [
                label for label in payment_label_list if re.match( specified_payment_label, label ) != None
            ]
            specified_category_main_items_items_dict = {} # 指定した 支出: カテゴリ

            time_list = []

            for line in lines[1:]: # 0 = columun 列 除外

                # get line data
                date = line[0] # 日付
                method = line[1] # 方法
                category_main = to_utf( line[2] ) # カテゴリ
                category_sub = to_utf( line[3] ) # 内訳
                payment_src = to_utf( line[4] ) # 支払元
                income_dst = to_utf( line[5] ) # 入金先
                income_price = int( line[10] ) # 収入
                payment_price = int( line[11] ) # 支出
                transfer_price = int( line[12] ) # 振替

                category_label = get_category_label( line ) # カテゴリ:内訳
                transfer_label = get_transfer_label( line ) # 振替: 支払元>入金先

                time_list.append( datetime_to_time( date_str_to_datetime( date ) ) )

                # カテゴリ別: 口座毎の支出
                if not category_items_dict.has_key( category_main ):
                    init_list = [ 0.0 ] * len( payment_label_list )
                    category_items_dict[ category_main ] = init_list

                index = payment_label_list.index( payment_src )
                category_items_dict[ category_main ][ index ] += payment_price


                # カテゴリ/内訳別: 口座毎の支出
                if not category_main_items_items_dict.has_key( category_main ):
                    category_main_items_items_dict[ category_main ] = {}

                if not category_main_items_items_dict[ category_main ].has_key( category_sub ):
                    init_list = [ 0.0 ] * len( payment_label_list )
                    category_main_items_items_dict[ category_main ][ category_sub ] = init_list

                index = payment_label_list.index( payment_src )
                category_main_items_items_dict[ category_main ][ category_sub ][ index ] += payment_price



                # 指定した カテゴリ/内訳別: 口座毎の支出
                if category_main in specified_category_main_list :
                    if not specified_category_main_items_items_dict.has_key( category_main ):
                        specified_category_main_items_items_dict[ category_main ] = {}

                    if category_sub in specified_category_sub_list :
                        if not specified_category_main_items_items_dict[ category_main ].has_key( category_sub ):
                            init_list = [ 0.0 ] * len( specified_payment_label_list )
                            specified_category_main_items_items_dict[ category_main ][ category_sub ] = init_list

                        if payment_src in specified_payment_label_list:
                            index = specified_payment_label_list.index( payment_src )
                            specified_category_main_items_items_dict[ category_main ][ category_sub ][ index ] += payment_price



                # 口座別: カテゴリ毎の支出
                if not payment_src_items_dict.has_key( payment_src ):
                    init_list = [ 0.0 ] * len( category_label_list )
                    payment_src_items_dict[ payment_src ] = init_list

                index = category_label_list.index( category_label )
                payment_src_items_dict[ payment_src ][ index ] += payment_price

                # 口座別: カテゴリ毎の収入
                if not income_dst_items_dict.has_key( income_dst ):
                    init_list = [ 0.0 ] * len( category_label_list )
                    income_dst_items_dict[ income_dst ] = init_list

                index = category_label_list.index( category_label )
                income_dst_items_dict[ income_dst ][ index ] += income_price


                # 口座別: 振替
                if not transfer_items_dict.has_key( transfer_label ):
                    init_list = [ 0.0 ] * len( transfer_label_list )
                    transfer_items_dict[ transfer_label ] = init_list

                index = transfer_label_list.index( transfer_label )
                transfer_items_dict[ transfer_label ][ index ] += transfer_price


            time_list = sorted( time_list )

            """ Chart ごとに整理 """

            # カテゴリの金額順に
            category_main_key_item_list = []
            for category_main_key in category_main_items_items_dict:

                total_payment = 0.0
                for category_sub_key in category_main_items_items_dict[ category_main_key ]:
                    payment_list = category_main_items_items_dict[ category_main_key ][ category_sub_key ]
                    total_payment += sum( payment_list )

                category_main_key_item_list.append( [ category_main_key, total_payment ] )

            category_main_key_item_list = sorted( category_main_key_item_list, key=itemgetter( 1 ) )
            category_main_key_list = [ item[0] for item in category_main_key_item_list ]
            category_main_key_list = sorted( category_main_key_list, reverse=1 ) # 金額が高い順に

            # カテゴリ ごとに分ける
            category_main_data_list = []
            for category_main_key in category_main_key_list:
                category_main_data_list.append(
                    [
                        category_main_key,
                        payment_label_list,
                        category_main_items_items_dict[ category_main_key ],
                        300
                    ]
                )


            # 指定 カテゴリの金額順に
            specified_category_main_key_item_list = []
            for specified_category_main_key in specified_category_main_items_items_dict:

                total_payment = 0.0
                for category_sub_key in specified_category_main_items_items_dict[ specified_category_main_key ]:
                    payment_list = specified_category_main_items_items_dict[ specified_category_main_key ][ category_sub_key ]
                    total_payment += sum( payment_list )

                specified_category_main_key_item_list.append( [ specified_category_main_key, total_payment ] )

            specified_category_main_key_item_list = sorted( specified_category_main_key_item_list, key=itemgetter( 1 ) )
            specified_category_main_key_list = [ item[0] for item in specified_category_main_key_item_list ]
            specified_category_main_key_list = sorted( specified_category_main_key_list, reverse=1 ) # 金額が高い順に

            # カテゴリ ごとに分ける
            specified_category_main_data_list = []
            for specified_category_main_key in specified_category_main_key_list:
                specified_category_main_data_list.append(
                    [
                        u'{0}:{1}'.format( specified_payment_label, u'+'.join( specified_category_sub_list ) ),
                        specified_payment_label_list,
                        specified_category_main_items_items_dict[ specified_category_main_key ],
                        300
                    ]
                )
            #print specified_category_main_data_list


            # Chart script line の作成
            div_line_list = []
            script_line_list = []

            for id, item in enumerate(
                [
                    [ u'カテゴリ別: 支出', payment_label_list, category_items_dict, 600 ]
                ] + specified_category_main_data_list + category_main_data_list + [
                    [ u'口座別: 支出', category_label_list, payment_src_items_dict, 600 ],
                    [ u'口座別: 収入', category_label_list, income_dst_items_dict, 600 ],
                    [ u'口座別: 振替', transfer_label_list, transfer_items_dict, 600 ]
                ]
            ):

                input_first_label = item[0]
                input_label_list = item[1]
                input_items_dict = item[2]
                height = item[3]

                label_list = [ input_first_label ] # column の先頭に追加
                label_list.extend( input_label_list )

                data_row_list = []
                data_row_list.append(
                    u"        [{0},{{ role: 'annotation' }}]".format(
                        ','.join( u"'{0}'".format( item ) for item in label_list )
                    )
                )

                items_dict = copy.copy( input_items_dict )

                src_key_list = items_dict.keys()
                src_key_list = sorted( src_key_list )

                # 項目の金額順に
                src_key_item_list = []
                for src_key in src_key_list:
                    payment_list = items_dict[ src_key ]
                    src_key_item_list.append( [ src_key, sum( payment_list ) ] )

                src_key_item_list = sorted( src_key_item_list, key=itemgetter( 1 ) )
                src_key_list = [ item[0] for item in src_key_item_list ]

                total_payment = 0.0

                for src_key in src_key_list:

                    payment_list = items_dict[ src_key ]

                    if sum( payment_list ) == 0.0: # No payment
                        continue

                    row_item_list = [ u'{0}'.format( item ) for item in payment_list ]
                    row_item_list[0:0] = [ u"'{0}'".format( src_key ) ] # column の先頭に追加
                    # 金額 row
                    data_row_list.append( u"        [{0},'¥{1}']".format(
                        ','.join( row_item_list ), int( sum( payment_list ) )
                    ) )

                    total_payment += sum( payment_list )

                if len( data_row_list ) <= 1: # column row しかない場合
                    continue

                script_line_list.append( get_html_chart_script_line_tmp().format(
                    **{
                        'title': u'{0} = ¥{1}'.format( input_first_label, int(total_payment) ),
                        'datas': u',\n'.join( data_row_list ),
                        'id': u'data{0:02d}'.format( id ),
                        'width': 1800,
                        'height': height,
                    }
                ) )

                """
    <Table>
        <tr>
            <td><div id="A_week" style="width: 900px; height: 140px;"></div></td>
            <td><div id="A_day" style="width: 900px; height: 140px;"></div></td>
        </tr>
    </Table>
                """

                div_line_list.append( get_html_div_line_tmp().format(
                    **{ 'id': u'data{0:02d}'.format( id ) }
                ) )

            """ html """
            html_line = get_html_line_tmp().format(
                **{
                    'date': u'{0}-{1}'.format(
                        src_time_to_str( min(time_list) ), src_time_to_str( max(time_list) )
                    ),
                    'script_lines': '\n'.join( script_line_list ),
                    'div_lines': '\n'.join( div_line_list ),
                }
            )

            basename = 'Zaim_{0}'.format(
                src_time_to_str( min(time_list), '{year}_{month}' ),
            )

            html_file_path = os.path.abspath( '{0}.html'.format( basename ) )

            with codecs.open( html_file_path, 'w', 'utf-8' ) as f:
                f.write( html_line )

            os.startfile( html_file_path )

    except:
        print traceback.format_exc()
        raw_input( '--- end ---' )
