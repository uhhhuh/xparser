#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import unittest
import xparse
import openpyxl
import pickle

class TestHell(unittest.TestCase):

    workbook = openpyxl.load_workbook('tests/test_data/test_book.xlsx', read_only=False)
    ws_name = workbook.sheetnames[0]
    xparse.ws = workbook[ws_name]
    
    def setUp(self):
        self.ws = xparse.ws

    def tearDown(self):
        del self.ws

            
    def test_open_xls(self):
        self.assertTrue(isinstance(self.ws, openpyxl.worksheet.worksheet.Worksheet))


    # General helpers
    
    def test_validate_dimensions(self):
        self.assertIs(xparse.validate_dimensions('A2:A999'), 'valid')
        self.assertIs(xparse.validate_dimensions('A225:X356'), 'valid')
        self.assertIs(xparse.validate_dimensions('A:A999'), 'invalid')
        self.assertIs(xparse.validate_dimensions('A1 :A2'), 'invalid')

    def test_not_empty(self):
        self.assertFalse(xparse.not_empty('-'))
        self.assertFalse(xparse.not_empty(' -  '))
        self.assertFalse(xparse.not_empty('не имеет'))
        self.assertFalse(xparse.not_empty('  не   имеет   '))
        self.assertFalse(xparse.not_empty(''))
        self.assertFalse(xparse.not_empty(' '))
        self.assertFalse(xparse.not_empty('   '))
        self.assertFalse(xparse.not_empty(None))
        self.assertFalse(xparse.not_empty('Автомобиль легковой:'))
        self.assertFalse(xparse.not_empty('Автоприцеп:'))
        self.assertFalse(xparse.not_empty('автомобили легковые:'))
        self.assertFalse(xparse.not_empty('иные транспортные средства:'))
        self.assertTrue(xparse.not_empty('опель'))
        self.assertTrue(xparse.not_empty('автоприцеп радуга'))
        self.assertTrue(xparse.not_empty('автоприцеп: радуга 5'))
        self.assertTrue(xparse.not_empty('Автомобиль легковой хонда'))
        self.assertTrue(xparse.not_empty('Автомобиль легковой: хонда'))
        self.assertTrue(xparse.not_empty(200.00001))
        self.assertTrue(xparse.not_empty(123))
        self.assertTrue(xparse.not_empty('опель-астра'))
        self.assertTrue(xparse.not_empty('ниссан - гтр'))
        
    """
    def test_get_sorted_coord(self):
        pass
    """


    # Main part

    def test_parse_person(self):
        with open('tests/test_data/data_all.pkl', 'rb') as pkl:
            self.data_all = pickle.load(pkl)
            
        self.assertListEqual(self.data_all, xparse.parse_person('A2:A787'))


    def test_check_lists_mismatch(self):
        list_a = [1,2,3,4,5,6,7,8,9,10]
        list_b = [1,2,3,4,5,6,8,9,10]
        self.assertEqual(xparse.check_lists_mismatch(list_a, list_b), (7,8))

        list_b = [1,3,4,5,6,7,8,9,10]
        self.assertEqual(xparse.check_lists_mismatch(list_a, list_b), (2,3))

        list_b = [1,2,3,4,5,6,7,8,9,10,11]
        self.assertEqual(xparse.check_lists_mismatch(list_a, list_b), (None, 11))


    def test_get_slot(self):
        '''testing against pickled slots_data from same file'''
        with open('tests/test_data/slots_data.pkl', 'rb') as pkl:
            self.slots_data = pickle.load(pkl)

        self.assertListEqual(self.slots_data, xparse.get_slot('A2', 'A787'))

    
    def test_shift_col(self):
        self.assertEqual(xparse.shift_col('A1'), 'B1')
        self.assertEqual(xparse.shift_col('M99'), 'N99')
        self.assertEqual(xparse.shift_col('D11', 3), 'G11')
        self.assertIsNone(xparse.shift_col('Z4', 2))
        self.assertEqual(xparse.shift_col('D9', 10), 'N9')
    

    def test_parse_ownership(self):
        with open('tests/test_data/person_slot.pkl', 'rb') as pkl:
            self.per_slot = pickle.load(pkl)
            
        with open('tests/test_data/parsed_ownership.pkl', 'rb') as pkl:
            self.parsed_own = pickle.load(pkl)
            
        self.assertListEqual(self.parsed_own, xparse.parse_ownership(self.per_slot))

        
    def test_parse_usage(self):
        with open('tests/test_data/person_slot.pkl', 'rb') as pkl:
            self.person_slot = pickle.load(pkl)
            
        with open('tests/test_data/parsed_usage.pkl', 'rb') as pkl:
            self.parsed_usage = pickle.load(pkl)
            
        self.assertListEqual(self.parsed_usage, xparse.parse_usage(self.person_slot))


    def test_parse_vehicle(self):
        with open('tests/test_data/person_slot.pkl', 'rb') as pkl:
            self.pers_slot = pickle.load(pkl)
            
        with open('tests/test_data/parsed_vehicle.pkl', 'rb') as pkl:
            self.parsed_vehicle = pickle.load(pkl)

        self.assertListEqual(self.parsed_vehicle, xparse.parse_vehicle(self.pers_slot))
    

    # Save data

    def test_map_data(self):
        with open('tests/test_data/person_data.pkl', 'rb') as pkl:
            self.person_data = pickle.load(pkl)
            
        with open('tests/test_data/mapped_person.pkl', 'rb') as pkl:
            self.mapped_data = pickle.load(pkl)
            
        self.assertEqual(self.mapped_data, xparse.map_data(self.person_data))

        with open('tests/test_data/person_data2.pkl', 'rb') as pkl:
            self.person_data2 = pickle.load(pkl)
            
        with open('tests/test_data/mapped_person2.pkl', 'rb') as pkl:
            self.mapped_data2 = pickle.load(pkl)
            
        self.assertEqual(self.mapped_data2, xparse.map_data(self.person_data2))
        
        with open('tests/test_data/person_data3.pkl', 'rb') as pkl:
            self.person_data3 = pickle.load(pkl)
            
        with open('tests/test_data/mapped_person3.pkl', 'rb') as pkl:
            self.mapped_data3 = pickle.load(pkl)
            
        self.assertEqual(self.mapped_data3, xparse.map_data(self.person_data3))


    def test_value_from_dict(self):
        self.assertEqual(xparse.value_from_dict('квартира', 'objectType'), 7)
        self.assertEqual(xparse.value_from_dict('Гараж', 'objectType'), 17)
        self.assertEqual(xparse.value_from_dict('не определено', 'objectType'), 0)
        self.assertEqual(xparse.value_from_dict('супруга', 'relationType'), 2)
        self.assertEqual(xparse.value_from_dict('долевая', 'ownershipType'), 3)
        self.assertEqual(xparse.value_from_dict('Грузия', 'country'), 2)
        self.assertEqual(xparse.value_from_dict('Кафиристан', 'country'), 'Кафиристан')
        self.assertEqual(xparse.value_from_dict('не определено', 'country'), 0)
        #self.assertEqual(xparse.value_from_dict('не имеет'), None)
        self.assertEqual(xparse.value_from_dict(' 0 '), None)
        # write tests for UnicodeDecodeError
        self.assertEqual(xparse.value_from_dict('  -  ', 'none_values'), None)
        self.assertEqual(xparse.value_from_dict(' -  ', 'none_values'), None)
        self.assertEqual(xparse.value_from_dict('-  ', 'none_values'), None)
        self.assertEqual(xparse.value_from_dict('  - ', 'none_values'), None)
        self.assertEqual(xparse.value_from_dict('  -', 'none_values'), None)

    
    def test_set_name(self):
        self.pd = {'name':'Иоганн Бах','relativeOf':'Амброзий Бах'}
        self.assertIsNone(xparse.set_name(self.pd))
        self.pd = {'name':'И Бах','relativeOf':None}
        self.assertEqual(xparse.set_name(self.pd), 'И Бах')
        self.pd = {'name':'супруг','relativeOf':None,'p':'1', 'start':'A1'}
        self.assertIsNone(xparse.set_name(self.pd))
        self.pd = {'name':'супруга','relativeOf':None,'p':'2', 'start':'A2'}
        self.assertIsNone(xparse.set_name(self.pd))
        self.pd = {'name':'несовершеннолетний ребёнок','relativeOf':None,
                            'p':'3', 'start':'A3'}
        self.assertIsNone(xparse.set_name(self.pd))
        self.pd = {'name':'несовершеннолетний ребенок','relativeOf':None,
                            'p':'4', 'start':'A4'}
        self.assertIsNone(xparse.set_name(self.pd))

                
    def test_set_position(self):
        self.pd = {'name':'Иоганн Бах','relativeOf':'Амброзий Бах'}
        self.assertIsNone(xparse.set_position(self.pd))
        self.pd = {'name':'Иоганн Бах','relativeOf':None,'position':'organist'}
        self.assertEqual(xparse.set_position(self.pd), 'organist')

    
    def test_set_ownership(self):
        self.pd = {'own_type':'индивидуальная'}
        self.assertEqual(xparse.set_ownership(self.pd), ('индивидуальная', None))

        self.pd = {'own_type':'2/3 доли'}
        self.assertEqual(xparse.set_ownership(self.pd), ('долевая', '2/3'))

        self.pd = {'own_type':'доли 361,2 балло-гектар'}
        self.assertEqual(xparse.set_ownership(self.pd), ('долевая', '361,2'))

        self.pd = {'own_type':'1/2 долевая'}
        self.assertEqual(xparse.set_ownership(self.pd), ('долевая', '1/2'))

        self.pd = {'own_type':'(321/421 доли)'}
        self.assertEqual(xparse.set_ownership(self.pd), ('долевая', '321/421'))

        self.pd = {'own_type':'совместная'}
        self.assertEqual(xparse.set_ownership(self.pd), ('совместная', None))

        self.pd = {'own_type':'массовая'}
        self.assertEqual(xparse.set_ownership(self.pd), ('массовая', None))


    def test_get_ownpart_amount(self):
        self.o = [
            ('(1/215 доли)',           '1/215'),
            ('(120/13185 доли)',       '120/13185'),
            ('(2/1261 доли)',          '2/1261'),
            ('(37/100 доли)',          '37/100'),
            ('(46/2922 доли)',         '46/2922'),
            ('(5/9 доли)',             '5/9'),
            ('(774/33017 доли)',       '774/33017'),
            ('(9/32 доли)',            '9/32'),
            ('(доля в праве 4,62 га)', '4,62'),
            ('1/ 3 доли',              '1/3'),
            ('долевая1 / 3 доли',      '1/3'),
            ('долевая1 /. 3 доли',     '1'),
            ('долевая 1 /3 доли',      '1/3'),
            ('долевая 1 / 3 доли',     '1/3'),
            ('долевая 1  /  3 доли',   '1'),
            ('долевая1/9доли',         '1/9'),            
            ('1/130 доли',             '1/130'),
            ('1/2 долевая',            '1/2'),
            ('1/234',                  '1/234'),
            ('1/2доли',                '1/2'),
            ('4/5 доли',               '4/5'),
            ('596/47200 доли',         '596/47200'),
            ('9/10 доли',              '9/10'),
            ('долевая',                'долевая'),
            ('долевая 3/5',            '3/5'),
            ('доли 1/16',              '1/16'),
            ('доли 2/3',               '2/3'),
            ('доли 361,2 балло-гектар','361,2'),
            ('доли 639/100000',        '639/100000'),
            ('индивидальная',          'индивидальная'),
            ('индивидуальная',         'индивидуальная'),
            ('индиивидуальная',        'индиивидуальная'),
            ('совместная',             'совместная'),
            ('срвместная',             'срвместная')]

        for n in range(len(self.o)):
            self.assertEqual(xparse.get_ownpart_amount(self.o[n][0]), self.o[n][1])


    def test_set_income(self):
        self.pd = {'income':'не имеет'}
        self.assertIsNone(xparse.set_income(self.pd))
        
        self.pd = {'income':' - '}
        self.assertIsNone(xparse.set_income(self.pd))

        self.pd = {'income':' '}
        self.assertIsNone(xparse.set_income(self.pd))

        self.pd = {'income':None}
        self.assertIsNone(xparse.set_income(self.pd))
        
        self.pd = {'income':'1031691.85'}
        self.assertEqual(xparse.set_income(self.pd), '1031691.85')
        
        self.pd = {'income':'не имеет '}
        self.assertIsNone(xparse.set_income(self.pd))
        
        self.pd = {'income':'941951'}
        self.assertEqual(xparse.set_income(self.pd), '941951')
        

    # Add data

    def test_make_blocks(self):
        with open('tests/test_data/blocks_by_p.pkl', 'rb') as pkl:
            self.blocks_by_p = pickle.load(pkl)

        with open('tests/test_data/data_all.pkl', 'rb') as pkl:
            self.data_all = pickle.load(pkl)

        self.assertEqual(self.blocks_by_p, xparse.make_blocks(self.data_all))


    def test_get_p(self):
        with open('tests/test_data/blocks.pkl', 'rb') as pkl:
            self.blocks = pickle.load(pkl)

        with open('tests/test_data/data_all.pkl', 'rb') as pkl:
            self.data_all = pickle.load(pkl)
            
        self.assertEqual(self.blocks, xparse.get_p(self.data_all))


    def test_set_relations(self):
        with open('tests/test_data/blocks_by_p_unrelated.pkl', 'rb') as pkl:
            self.unrelated = pickle.load(pkl)

        with open('tests/test_data/blocks_by_p_related.pkl', 'rb') as pkl:
            self.related = pickle.load(pkl)
            
        xparse.set_relations(self.unrelated)
        self.assertEqual(self.related, self.related)

        
    # Pre save
    
    def test_parent_to_child(self):
        self.assertEqual(xparse.parent_to_child('realties'), 'realty')
        self.assertEqual(xparse.parent_to_child('transports'), 'transport')
        self.assertEqual(xparse.parent_to_child('persons'), 'person')

    
    def test_load_file(self):
        self.xlsx = xparse.load_file('data/book_100.xlsx')
        self.assertTrue(isinstance(self.xlsx, openpyxl.worksheet.worksheet.Worksheet))
        self.assertEqual(self.xlsx.title, 'Sheet1')
        self.assertEqual(self.xlsx.calculate_dimension(), 'A1:AMJ787')
