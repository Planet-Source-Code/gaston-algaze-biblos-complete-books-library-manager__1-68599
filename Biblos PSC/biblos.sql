# SQL Manager 2005 for MySQL 3.7.0.1
# ---------------------------------------
# Host     : localhost
# Port     : 3306
# Database : biblos


SET FOREIGN_KEY_CHECKS=0;

DROP DATABASE IF EXISTS `biblos`;

CREATE DATABASE `biblos`
    CHARACTER SET 'latin1'
    COLLATE 'latin1_swedish_ci';

#
# Structure for the `campos` table : 
#

DROP TABLE IF EXISTS `campos`;

CREATE TABLE `campos` (
  `id_campo` int(11) NOT NULL auto_increment,
  `id_tabla` int(11) NOT NULL,
  `nombre` varchar(20) default NULL,
  `tipo_datos` varchar(20) default NULL,
  `fecha_creacion` date default NULL,
  PRIMARY KEY  (`id_campo`,`id_tabla`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `categorias` table : 
#

DROP TABLE IF EXISTS `categorias`;

CREATE TABLE `categorias` (
  `id_categoria` int(11) NOT NULL auto_increment,
  `descripcion` varchar(200) default NULL,
  `fecha_alta` date default NULL,
  PRIMARY KEY  (`id_categoria`),
  CONSTRAINT `categorias_fk` FOREIGN KEY (`id_categoria`) REFERENCES `categorias` (`id_categoria`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `categorias_subcategorias` table : 
#

DROP TABLE IF EXISTS `categorias_subcategorias`;

CREATE TABLE `categorias_subcategorias` (
  `id_categoria` int(11) NOT NULL,
  `id_subcategoria` int(11) NOT NULL,
  PRIMARY KEY  (`id_categoria`,`id_subcategoria`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `editoriales` table : 
#

DROP TABLE IF EXISTS `editoriales`;

CREATE TABLE `editoriales` (
  `id_editorial` int(11) NOT NULL auto_increment,
  `nombre` varchar(255) default NULL,
  `tel_1` varchar(16) default NULL,
  `tel_2` varchar(16) default NULL,
  `email` varchar(100) default NULL,
  `web` varchar(100) default NULL,
  `domicilio_calle` varchar(255) default NULL,
  `domicilio_piso` varchar(10) default NULL,
  `domicilio_nro` varchar(10) default NULL,
  `domicilio_cod_postal` varchar(12) default NULL,
  `fecha_alta` date default NULL,
  `fecha_baja` date default NULL,
  PRIMARY KEY  (`id_editorial`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `fichas` table : 
#

DROP TABLE IF EXISTS `fichas`;

CREATE TABLE `fichas` (
  `id_ficha` int(11) NOT NULL auto_increment,
  `id_usuario` int(11) NOT NULL,
  `titulo` varchar(100) default NULL,
  PRIMARY KEY  (`id_ficha`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `funcion_tabla_permiso` table : 
#

DROP TABLE IF EXISTS `funcion_tabla_permiso`;

CREATE TABLE `funcion_tabla_permiso` (
  `id_funcion` int(11) NOT NULL,
  `id_tabla` int(11) NOT NULL,
  `id_permiso` int(11) NOT NULL,
  `fecha_creacion` date default NULL,
  PRIMARY KEY  (`id_funcion`,`id_tabla`,`id_permiso`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `funciones` table : 
#

DROP TABLE IF EXISTS `funciones`;

CREATE TABLE `funciones` (
  `id_funcion` int(11) NOT NULL auto_increment,
  `nombre` varchar(100) default NULL,
  `fecha_creacion` date default NULL,
  PRIMARY KEY  (`id_funcion`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `libros` table : 
#

DROP TABLE IF EXISTS `libros`;

CREATE TABLE `libros` (
  `id_libro` int(11) NOT NULL auto_increment,
  `codigo_libro` varchar(30) default NULL,
  `titulo` varchar(255) default NULL,
  `autor` varchar(255) default NULL,
  `ISBN` varchar(20) default NULL,
  `anio` char(4) default NULL,
  `fecha_alta` date default NULL,
  `id_editorial` int(11) default NULL,
  `id_ubicacion` int(11) default NULL,
  `id_tipo_material` int(11) default NULL,
  `fecha_baja` date default NULL,
  PRIMARY KEY  (`id_libro`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `libros_categorias_subcategorias` table : 
#

DROP TABLE IF EXISTS `libros_categorias_subcategorias`;

CREATE TABLE `libros_categorias_subcategorias` (
  `id_categoria` int(11) NOT NULL,
  `id_subcategoria` int(11) NOT NULL,
  `id_libro` int(11) NOT NULL,
  `fecha_creacion` date default NULL,
  PRIMARY KEY  (`id_categoria`,`id_subcategoria`,`id_libro`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `links` table : 
#

DROP TABLE IF EXISTS `links`;

CREATE TABLE `links` (
  `id_link` int(11) NOT NULL auto_increment,
  `id_ficha` int(11) NOT NULL,
  `descripcion` varchar(255) default NULL,
  `direccion` varchar(255) default NULL,
  PRIMARY KEY  (`id_link`,`id_ficha`),
  UNIQUE KEY `id_link` (`id_link`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `operaciones` table : 
#

DROP TABLE IF EXISTS `operaciones`;

CREATE TABLE `operaciones` (
  `id_operacion` int(11) NOT NULL auto_increment,
  `descripcion` int(11) default NULL,
  PRIMARY KEY  (`id_operacion`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `permisos` table : 
#

DROP TABLE IF EXISTS `permisos`;

CREATE TABLE `permisos` (
  `id_permiso` int(11) NOT NULL auto_increment,
  `descripcion` varchar(20) default NULL,
  PRIMARY KEY  (`id_permiso`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `prestamos` table : 
#

DROP TABLE IF EXISTS `prestamos`;

CREATE TABLE `prestamos` (
  `id_prestamo` int(11) NOT NULL auto_increment,
  `fecha_desde` date default NULL,
  `fecha_hasta` date default NULL,
  `id_usuario` int(11) default NULL,
  `id_bibliotecaria` int(11) default NULL,
  `id_libro` int(11) default NULL,
  `fecha_devolucion` date default NULL,
  PRIMARY KEY  (`id_prestamo`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `restricciones` table : 
#

DROP TABLE IF EXISTS `restricciones`;

CREATE TABLE `restricciones` (
  `id_restriccion` int(11) NOT NULL auto_increment,
  `id_campo` int(11) NOT NULL,
  `id_tabla` int(11) NOT NULL,
  `id_operacion` int(11) default NULL,
  `valor` varchar(255) default NULL,
  PRIMARY KEY  (`id_restriccion`,`id_campo`,`id_tabla`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `roles` table : 
#

DROP TABLE IF EXISTS `roles`;

CREATE TABLE `roles` (
  `id_rol` int(11) NOT NULL auto_increment,
  `descripcion` varchar(100) default NULL,
  `fecha_creacion` date default NULL,
  PRIMARY KEY  (`id_rol`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `roles_funciones` table : 
#

DROP TABLE IF EXISTS `roles_funciones`;

CREATE TABLE `roles_funciones` (
  `id_rol` int(11) NOT NULL,
  `id_funcion` int(11) NOT NULL,
  PRIMARY KEY  (`id_rol`,`id_funcion`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `roles_restricciones` table : 
#

DROP TABLE IF EXISTS `roles_restricciones`;

CREATE TABLE `roles_restricciones` (
  `id_rol` int(11) NOT NULL,
  `id_restriccion` int(11) NOT NULL,
  PRIMARY KEY  (`id_rol`,`id_restriccion`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `roles_usuarios` table : 
#

DROP TABLE IF EXISTS `roles_usuarios`;

CREATE TABLE `roles_usuarios` (
  `id_usuario` int(11) NOT NULL,
  `id_rol` int(11) NOT NULL,
  `fecha_creacion` date default NULL,
  PRIMARY KEY  (`id_usuario`,`id_rol`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `subcategorias` table : 
#

DROP TABLE IF EXISTS `subcategorias`;

CREATE TABLE `subcategorias` (
  `id_subcategoria` int(11) NOT NULL auto_increment,
  `descripcion` varchar(100) default NULL,
  `fecha_alta` date default NULL,
  PRIMARY KEY  (`id_subcategoria`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tablas` table : 
#

DROP TABLE IF EXISTS `tablas`;

CREATE TABLE `tablas` (
  `id_tabla` int(11) NOT NULL auto_increment,
  `nombre` varchar(100) default NULL,
  `fecha_creacion` date default NULL,
  PRIMARY KEY  (`id_tabla`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `tipo_material` table : 
#

DROP TABLE IF EXISTS `tipo_material`;

CREATE TABLE `tipo_material` (
  `id_tipo_material` int(11) NOT NULL auto_increment,
  `decripcion` varchar(255) default NULL,
  PRIMARY KEY  (`id_tipo_material`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `ubicaciones` table : 
#

DROP TABLE IF EXISTS `ubicaciones`;

CREATE TABLE `ubicaciones` (
  `id_ubicacion` int(11) NOT NULL auto_increment,
  `descripcion` varchar(20) default NULL,
  `titulo` varchar(255) default NULL,
  PRIMARY KEY  (`id_ubicacion`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Structure for the `usuarios` table : 
#

DROP TABLE IF EXISTS `usuarios`;

CREATE TABLE `usuarios` (
  `id_usuario` int(11) NOT NULL auto_increment,
  `nombre` varchar(255) default NULL,
  `apellido` varchar(255) default NULL,
  `dni` varchar(10) default NULL,
  `matricula` varchar(10) default NULL,
  `fecha_nacimiento` date default NULL,
  `domicilio_calle` varchar(255) default NULL,
  `domicilio_piso` varchar(10) default NULL,
  `domicilio_nro` varchar(10) default NULL,
  `domicilio_cod_postal` varchar(12) default NULL,
  `tel1` varchar(16) default NULL,
  `tel2` varchar(16) default NULL,
  `fecha_baja` date default NULL,
  PRIMARY KEY  (`id_usuario`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

