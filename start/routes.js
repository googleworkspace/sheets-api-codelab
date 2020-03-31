/*
  Copyright 2016 Google, Inc.

  Licensed to the Apache Software Foundation (ASF) under one or more contributor
  license agreements. See the NOTICE file distributed with this work for
  additional information regarding copyright ownership. The ASF licenses this
  file to you under the Apache License, Version 2.0 (the "License"); you may not
  use this file except in compliance with the License. You may obtain a copy of
  the License at

  http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
  WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
  License for the specific language governing permissions and limitations under
  the License.
*/

'use strict';

var express = require('express');
var router = express.Router();
var models = require('./models');
var Sequelize = require('sequelize');

// TODO: Show spreadsheets on the main page.
router.get('/', function(req, res, next) {
  var options = {
    order: [['createdAt', 'DESC']],
    raw: true
  };
  models.Order.findAll(options)
  .then(function(orders) {
    res.render('index', {
      orders: orders
    });
  }, function(err) {
    next(err);
  });
});

router.get('/create', function(req, res, next) {
  res.render('upsert');
});

router.get('/edit/:id', function(req, res, next) {
  models.Order.findByPk(req.params.id).then(function(order) {
    if (order) {
      res.render('upsert', {
        order: order.toJSON()
      });
    } else {
      next(new Error('Order not found: ' + req.params.id));
    }
  });
});

router.get('/delete/:id', function(req, res, next) {
  models.Order.findByPk(req.params.id)
    .then(function(order) {
      if (!order) {
        throw new Error('Order not found: ' + req.params.id);
      }
      return order.destroy();
    })
    .then(function() {
      res.redirect('/');
    }, function(err) {
      next(err);
    });
});

router.post('/upsert', function(req, res, next) {
  models.Order.upsert(req.body).then(function() {
    res.redirect('/');
  }, function(err) {
    next(err);
  });
});

// TODO: Add route for creating spreadsheet.



// TODO: Add route for syncing spreadsheet.



module.exports = router;
