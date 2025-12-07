#!/usr/bin/env python3
"""
SuperSlicer/PrusaSlicer to OrcaSlicer Profile Converter

Converts printer, print, and filament profile settings from PrusaSlicer
and SuperSlicer INI files to JSON format for use with OrcaSlicer.

Original Perl script by theophile:
https://github.com/theophile/SuperSlicer_to_Orca_scripts

Python conversion maintains feature parity with the original.
"""

import argparse
import json
import os
import platform
import re
import sys
import tempfile
from glob import glob
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Constants
ORCA_SLICER_VERSION = '1.6.0.0'

# Mapping of system directories by OS
SYSTEM_DIRECTORIES = {
    'os': {
        'Linux': ['.config'],
        'Windows': ['AppData', 'Roaming'],
        'Darwin': ['Library', 'Application Support'],
    },
    'output': {
        'filament': ['user', 'default', 'filament'],
        'print': ['user', 'default', 'process'],
        'printer': ['user', 'default', 'machine'],
    },
}

# Parameter mappings for translating source INI settings to OrcaSlicer JSON keys
PARAMETER_MAP = {
    'print': {
        'arc_fitting': 'enable_arc_fitting',
        'bottom_solid_layers': 'bottom_shell_layers',
        'bottom_solid_min_thickness': 'bottom_shell_thickness',
        'bridge_acceleration': 'bridge_acceleration',
        'bridge_angle': 'bridge_angle',
        'bridge_overlap_min': 'bridge_density',
        'dont_support_bridges': 'bridge_no_support',
        'bridge_speed_internal': 'internal_bridge_speed',
        'brim_ears': 'brim_ears',
        'brim_ears_detection_length': 'brim_ears_detection_length',
        'brim_ears_max_angle': 'brim_ears_max_angle',
        'brim_separation': 'brim_object_gap',
        'brim_width': 'brim_width',
        'brim_speed': 'skirt_speed',
        'compatible_printers_condition': 'compatible_printers_condition',
        'compatible_printers': 'compatible_printers',
        'default_acceleration': 'default_acceleration',
        'overhangs': 'detect_overhang_wall',
        'thin_walls': 'detect_thin_wall',
        'draft_shield': 'draft_shield',
        'first_layer_size_compensation': 'elefant_foot_compensation',
        'elefant_foot_compensation': 'elefant_foot_compensation',
        'enable_dynamic_overhang_speeds': 'enable_overhang_speed',
        'extra_perimeters_on_overhangs': 'extra_perimeters_on_overhangs',
        'extra_perimeters_odd_layers': 'alternate_extra_wall',
        'wipe_tower': 'enable_prime_tower',
        'wipe_speed': 'wipe_speed',
        'ensure_vertical_shell_thickness': 'ensure_vertical_shell_thickness',
        'gap_fill_min_length': 'filter_out_gap_fill',
        'gcode_comments': 'gcode_comments',
        'gcode_label_objects': 'gcode_label_objects',
        'machine_limits_usage': 'emit_machine_limits_to_gcode',
        'infill_anchor_max': 'infill_anchor_max',
        'infill_anchor': 'infill_anchor',
        'fill_angle': 'infill_direction',
        'infill_overlap': 'infill_wall_overlap',
        'infill_first': 'is_infill_first',
        'inherits': 'inherits',
        'extrusion_width': 'line_width',
        'extrusion_multiplier': 'print_flow_ratio',
        'first_layer_acceleration': 'initial_layer_acceleration',
        'first_layer_extrusion_width': 'initial_layer_line_width',
        'first_layer_height': 'initial_layer_print_height',
        'interface_shells': 'interface_shells',
        'perimeter_extrusion_width': 'inner_wall_line_width',
        'seam_gap': 'seam_gap',
        'solid_infill_acceleration': 'internal_solid_infill_acceleration',
        'solid_infill_extrusion_width': 'internal_solid_infill_line_width',
        'ironing_flowrate': 'ironing_flow',
        'ironing_spacing': 'ironing_spacing',
        'ironing_speed': 'ironing_speed',
        'layer_height': 'layer_height',
        'init_z_rotate': 'preferred_orientation',
        'spiral_vase': 'spiral_mode',
        'solid_infill_extruder': 'solid_infill_filament',
        'support_material_extruder': 'support_filament',
        'infill_extruder': 'sparse_infill_filament',
        'perimeter_extruder': 'wall_filament',
        'first_layer_extruder': 'first_layer_filament',
        'support_material_interface_extruder': 'support_interface_filament',
        'avoid_crossing_perimeters_max_detour': 'max_travel_detour_distance',
        'min_bead_width': 'min_bead_width',
        'min_feature_size': 'min_feature_size',
        'solid_infill_below_area': 'minimum_sparse_infill_area',
        'only_one_perimeter_first_layer': 'only_one_wall_first_layer',
        'only_one_perimeter_top': 'only_one_wall_top',
        'ooze_prevention': 'ooze_prevention',
        'extra_perimeters_overhangs': 'extra_perimeters_on_overhangs',
        'overhangs_reverse': 'overhang_reverse',
        'overhangs_reverse_threshold': 'overhang_reverse_threshold',
        'perimeter_acceleration': 'inner_wall_acceleration',
        'external_perimeter_acceleration': 'outer_wall_acceleration',
        'external_perimeter_extrusion_width': 'outer_wall_line_width',
        'post_process': 'post_process',
        'wipe_tower_brim_width': 'prime_tower_brim_width',
        'wipe_tower_width': 'prime_tower_width',
        'raft_contact_distance': 'raft_contact_distance',
        'raft_expansion': 'raft_expansion',
        'raft_first_layer_density': 'raft_first_layer_density',
        'raft_first_layer_expansion': 'raft_first_layer_expansion',
        'raft_layers': 'raft_layers',
        'avoid_crossing_perimeters': 'reduce_crossing_wall',
        'only_retract_when_crossing_perimeters': 'reduce_infill_retraction',
        'resolution': 'resolution',
        'seam_position': 'seam_position',
        'skirt_distance': 'skirt_distance',
        'skirt_height': 'skirt_height',
        'skirts': 'skirt_loops',
        'slice_closing_radius': 'slice_closing_radius',
        'slicing_mode': 'slicing_mode',
        'small_perimeter_min_length': 'small_perimeter_threshold',
        'infill_acceleration': 'sparse_infill_acceleration',
        'fill_density': 'sparse_infill_density',
        'infill_extrusion_width': 'sparse_infill_line_width',
        'staggered_inner_seams': 'staggered_inner_seams',
        'standby_temperature_delta': 'standby_temperature_delta',
        'hole_to_polyhole': 'hole_to_polyhole',
        'hole_to_polyhole_threshold': 'hole_to_polyhole_threshold',
        'hole_to_polyhole_twisted': 'hole_to_polyhole_twisted',
        'support_material': 'enable_support',
        'support_material_angle': 'support_angle',
        'support_material_enforce_layers': 'enforce_support_layers',
        'support_material_spacing': 'support_base_pattern_spacing',
        'support_material_contact_distance': 'support_top_z_distance',
        'first_layer_size_compensation_layers': 'elefant_foot_compensation_layers',
        'support_material_bottom_contact_distance': 'support_bottom_z_distance',
        'support_material_bottom_interface_layers': 'support_interface_bottom_layers',
        'support_material_interface_contact_loops': 'support_interface_loop_pattern',
        'support_material_interface_spacing': 'support_interface_spacing',
        'support_material_interface_layers': 'support_interface_top_layers',
        'support_material_extrusion_width': 'support_line_width',
        'support_material_buildplate_only': 'support_on_build_plate_only',
        'support_material_threshold': 'support_threshold_angle',
        'thick_bridges': 'thick_bridges',
        'top_solid_layers': 'top_shell_layers',
        'top_solid_min_thickness': 'top_shell_thickness',
        'top_solid_infill_acceleration': 'top_surface_acceleration',
        'top_infill_extrusion_width': 'top_surface_line_width',
        'min_width_top_surface': 'min_width_top_surface',
        'travel_acceleration': 'travel_acceleration',
        'travel_speed_z': 'travel_speed_z',
        'travel_speed': 'travel_speed',
        'support_tree_angle': 'tree_support_branch_angle',
        'support_tree_angle_slow': 'tree_support_angle_slow',
        'support_tree_branch_diameter': 'tree_support_branch_diameter',
        'support_tree_branch_diameter_angle': 'tree_support_branch_diameter_angle',
        'support_tree_branch_diameter_double_wall': 'tree_support_branch_diameter_double_wall',
        'support_tree_tip_diameter': 'tree_support_tip_diameter',
        'support_tree_top_rate': 'tree_support_top_rate',
        'wall_distribution_count': 'wall_distribution_count',
        'perimeter_generator': 'wall_generator',
        'perimeters': 'wall_loops',
        'wall_transition_angle': 'wall_transition_angle',
        'wall_transition_filter_deviation': 'wall_transition_filter_deviation',
        'wall_transition_length': 'wall_transition_length',
        'wipe_tower_no_sparse_layers': 'wipe_tower_no_sparse_layers',
        'xy_size_compensation': 'xy_contour_compensation',
        'z_offset': 'z_offset',
        'xy_inner_size_compensation': 'xy_hole_compensation',
        'support_material_layer_height': 'independent_support_layer_height',
        'fill_pattern': 'sparse_infill_pattern',
        'solid_fill_pattern': 'internal_solid_infill_pattern',
        'output_filename_format': 'filename_format',
        'support_material_pattern': 'support_base_pattern',
        'support_material_interface_pattern': 'support_interface_pattern',
        'top_fill_pattern': 'top_surface_pattern',
        'support_material_xy_spacing': 'support_object_xy_distance',
        'fuzzy_skin_point_dist': 'fuzzy_skin_point_distance',
        'fuzzy_skin_thickness': 'fuzzy_skin_thickness',
        'fuzzy_skin': 'fuzzy_skin',
        'bottom_fill_pattern': 'bottom_surface_pattern',
        'bridge_flow_ratio': 'bridge_flow',
        'fill_top_flow_ratio': 'top_solid_infill_flow_ratio',
        'first_layer_flow_ratio': 'bottom_solid_infill_flow_ratio',
        'infill_every_layers': 'infill_combination',
        'complete_objects': 'print_sequence',
        'brim_type': 'brim_type',
        'notes': 'notes',
        'support_material_style': 'support_style',
        'ironing': 'ironing',
        'ironing_type': 'ironing_type',
        'ironing_angle': 'ironing_angle',
        'external_perimeters_first': 'external_perimeters_first',
        'remaining_times': 'disable_m73',
    },
    'filament': {
        'bed_temperature': [
            'hot_plate_temp', 'cool_plate_temp',
            'eng_plate_temp', 'textured_plate_temp'
        ],
        'bridge_fan_speed': 'overhang_fan_speed',
        'chamber_temperature': 'chamber_temperature',
        'disable_fan_first_layers': 'close_fan_the_first_x_layers',
        'end_filament_gcode': 'filament_end_gcode',
        'external_perimeter_fan_speed': 'overhang_fan_threshold',
        'extrusion_multiplier': 'filament_flow_ratio',
        'fan_always_on': 'reduce_fan_stop_start_freq',
        'fan_below_layer_time': 'fan_cooling_layer_time',
        'fan_speedup_time': 'fan_speedup_time',
        'fan_speedup_overhangs': 'fan_speedup_overhangs',
        'fan_kickstart': 'fan_kickstart',
        'filament_colour': 'default_filament_colour',
        'filament_cost': 'filament_cost',
        'filament_density': 'filament_density',
        'filament_deretract_speed': 'filament_deretraction_speed',
        'filament_diameter': 'filament_diameter',
        'filament_max_volumetric_speed': 'filament_max_volumetric_speed',
        'filament_notes': 'filament_notes',
        'filament_retract_before_travel': 'filament_retraction_minimum_travel',
        'filament_retract_before_wipe': 'filament_retract_before_wipe',
        'filament_retract_layer_change': 'filament_retract_when_changing_layer',
        'filament_retract_length': 'filament_retraction_length',
        'filament_retract_lift': 'filament_z_hop',
        'filament_retract_lift_above': 'filament_retract_lift_above',
        'filament_retract_lift_below': 'filament_retract_lift_below',
        'filament_retract_restart_extra': 'filament_retract_restart_extra',
        'filament_retract_speed': 'filament_retraction_speed',
        'filament_shrink': 'filament_shrink',
        'filament_soluble': 'filament_soluble',
        'filament_type': 'filament_type',
        'filament_wipe': 'filament_wipe',
        'first_layer_bed_temperature': [
            'hot_plate_temp_initial_layer',
            'cool_plate_temp_initial_layer',
            'eng_plate_temp_initial_layer',
            'textured_plate_temp_initial_layer'
        ],
        'first_layer_temperature': 'nozzle_temperature_initial_layer',
        'full_fan_speed_layer': 'full_fan_speed_layer',
        'inherits': 'inherits',
        'max_fan_speed': 'fan_max_speed',
        'min_fan_speed': 'fan_min_speed',
        'min_print_speed': 'slow_down_min_speed',
        'slowdown_below_layer_time': 'slow_down_layer_time',
        'start_filament_gcode': 'filament_start_gcode',
        'support_material_interface_fan_speed': 'support_material_interface_fan_speed',
        'temperature': 'nozzle_temperature',
        'compatible_printers_condition': 'compatible_printers_condition',
        'compatible_printers': 'compatible_printers',
        'compatible_prints_condition': 'compatible_prints_condition',
        'compatible_prints': 'compatible_prints',
        'filament_vendor': 'filament_vendor',
        'filament_minimal_purge_on_wipe_tower': 'filament_minimal_purge_on_wipe_tower',
    },
    'printer': {
        'bed_custom_model': 'bed_custom_model',
        'bed_custom_texture': 'bed_custom_texture',
        'before_layer_gcode': 'before_layer_change_gcode',
        'toolchange_gcode': 'change_filament_gcode',
        'default_filament_profile': 'default_filament_profile',
        'default_print_profile': 'default_print_profile',
        'deretract_speed': 'deretraction_speed',
        'gcode_flavor': 'gcode_flavor',
        'inherits': 'inherits',
        'layer_gcode': 'layer_change_gcode',
        'feature_gcode': 'change_extrusion_role_gcode',
        'end_gcode': 'machine_end_gcode',
        'machine_max_acceleration_e': 'machine_max_acceleration_e',
        'machine_max_acceleration_extruding': 'machine_max_acceleration_extruding',
        'machine_max_acceleration_retracting': 'machine_max_acceleration_retracting',
        'machine_max_acceleration_travel': 'machine_max_acceleration_travel',
        'machine_max_acceleration_x': 'machine_max_acceleration_x',
        'machine_max_acceleration_y': 'machine_max_acceleration_y',
        'machine_max_acceleration_z': 'machine_max_acceleration_z',
        'machine_max_feedrate_e': 'machine_max_speed_e',
        'machine_max_feedrate_x': 'machine_max_speed_x',
        'machine_max_feedrate_y': 'machine_max_speed_y',
        'machine_max_feedrate_z': 'machine_max_speed_z',
        'machine_max_jerk_e': 'machine_max_jerk_e',
        'machine_max_jerk_x': 'machine_max_jerk_x',
        'machine_max_jerk_y': 'machine_max_jerk_y',
        'machine_max_jerk_z': 'machine_max_jerk_z',
        'machine_min_extruding_rate': 'machine_min_extruding_rate',
        'machine_min_travel_rate': 'machine_min_travel_rate',
        'pause_print_gcode': 'machine_pause_gcode',
        'start_gcode': 'machine_start_gcode',
        'max_layer_height': 'max_layer_height',
        'min_layer_height': 'min_layer_height',
        'nozzle_diameter': 'nozzle_diameter',
        'print_host': 'print_host',
        'printer_notes': 'printer_notes',
        'bed_shape': 'printable_area',
        'max_print_height': 'printable_height',
        'printer_technology': 'printer_technology',
        'printer_variant': 'printer_variant',
        'retract_before_wipe': 'retract_before_wipe',
        'retract_length_toolchange': 'retract_length_toolchange',
        'retract_restart_extra_toolchange': 'retract_restart_extra_toolchange',
        'retract_restart_extra': 'retract_restart_extra',
        'retract_layer_change': 'retract_when_changing_layer',
        'retract_length': 'retraction_length',
        'retract_lift': 'z_hop',
        'retract_lift_top': 'retract_lift_enforce',
        'retract_before_travel': 'retraction_minimum_travel',
        'retract_speed': 'retraction_speed',
        'silent_mode': 'silent_mode',
        'single_extruder_multi_material': 'single_extruder_multi_material',
        'thumbnails': 'thumbnails',
        'thumbnails_format': 'thumbnails_format',
        'template_custom_gcode': 'template_custom_gcode',
        'use_firmware_retraction': 'use_firmware_retraction',
        'use_relative_e_distances': 'use_relative_e_distances',
        'wipe': 'wipe',
    },
    'physical_printer': {
        'host_type': True,
        'print_host': True,
        'printer_technology': True,
        'printhost_apikey': True,
        'printhost_authorization_type': True,
        'printhost_cafile': True,
        'printhost_password': True,
        'printhost_port': True,
        'printhost_ssl_ignore_revoke': True,
        'printhost_user': True,
    },
}

# Printer parameters that may be comma-separated lists
MULTIVALUE_PARAMS = {
    'max_layer_height': 'single',
    'min_layer_height': 'single',
    'deretract_speed': 'single',
    'default_filament_profile': 'single',
    'machine_max_acceleration_e': 'array',
    'machine_max_acceleration_extruding': 'array',
    'machine_max_acceleration_retracting': 'array',
    'machine_max_acceleration_travel': 'array',
    'machine_max_acceleration_x': 'array',
    'machine_max_acceleration_y': 'array',
    'machine_max_acceleration_z': 'array',
    'machine_max_feedrate_e': 'array',
    'machine_max_feedrate_x': 'array',
    'machine_max_feedrate_y': 'array',
    'machine_max_feedrate_z': 'array',
    'machine_max_jerk_e': 'array',
    'machine_max_jerk_x': 'array',
    'machine_max_jerk_y': 'array',
    'machine_max_jerk_z': 'array',
    'machine_min_extruding_rate': 'array',
    'machine_min_travel_rate': 'array',
    'nozzle_diameter': 'single',
    'bed_shape': 'array',
    'retract_before_wipe': 'single',
    'retract_length_toolchange': 'single',
    'retract_restart_extra_toolchange': 'single',
    'retract_restart_extra': 'single',
    'retract_layer_change': 'single',
    'retract_length': 'single',
    'retract_lift': 'single',
    'retract_before_travel': 'single',
    'retract_speed': 'single',
    'thumbnails': 'array',
    'extruder_offset': 'single',
    'retract_lift_above': 'single',
    'retract_lift_below': 'single',
    'wipe': 'single',
}

# Filament type mappings
FILAMENT_TYPES = {
    'PET': 'PETG',
    'FLEX': 'TPU',
    'NYLON': 'PA',
}

# Default max volumetric speeds
DEFAULT_MVS = {
    'PLA': '15',
    'PET': '10',
    'ABS': '12',
    'ASA': '12',
    'FLEX': '3.2',
    'NYLON': '12',
    'PVA': '12',
    'PC': '12',
    'PSU': '8',
    'HIPS': '8',
    'EDGE': '8',
    'NGEN': '8',
    'PP': '8',
    'PEI': '8',
    'PEEK': '8',
    'PEKK': '8',
    'POM': '8',
    'PVDF': '8',
    'SCAFF': '8',
}

# Speed parameters mapping
SPEED_SEQUENCE = [
    'perimeter_speed', 'external_perimeter_speed',
    'solid_infill_speed', 'infill_speed',
    'small_perimeter_speed', 'top_solid_infill_speed',
    'gap_fill_speed', 'support_material_speed',
    'support_material_interface_speed', 'bridge_speed',
    'first_layer_speed', 'first_layer_infill_speed',
]

SPEED_PARAMS = {
    'perimeter_speed': 'inner_wall_speed',
    'external_perimeter_speed': 'outer_wall_speed',
    'small_perimeter_speed': 'small_perimeter_speed',
    'solid_infill_speed': 'internal_solid_infill_speed',
    'infill_speed': 'sparse_infill_speed',
    'top_solid_infill_speed': 'top_surface_speed',
    'gap_fill_speed': 'gap_infill_speed',
    'support_material_speed': 'support_speed',
    'support_material_interface_speed': 'support_interface_speed',
    'bridge_speed': 'bridge_speed',
    'first_layer_speed': 'initial_layer_speed',
    'first_layer_infill_speed': 'initial_layer_infill_speed',
}

# Seam position mappings
SEAM_POSITIONS = {
    'cost': 'nearest',
    'random': 'random',
    'allrandom': 'random',
    'aligned': 'aligned',
    'contiguous': 'aligned',
    'rear': 'back',
    'nearest': 'nearest',
}

# Infill type mappings
INFILL_TYPES = {
    '3dhoneycomb': '3dhoneycomb',
    'adaptivecubic': 'adaptivecubic',
    'alignedrectilinear': 'alignedrectilinear',
    'archimedeanchords': 'archimedeanchords',
    'concentric': 'concentric',
    'concentricgapfill': 'concentric',
    'cubic': 'cubic',
    'grid': 'grid',
    'gyroid': 'gyroid',
    'hilbertcurve': 'hilbertcurve',
    'honeycomb': 'honeycomb',
    'lightning': 'lightning',
    'line': 'line',
    'monotonic': 'monotonic',
    'monotonicgapfill': 'monotonic',
    'monotoniclines': 'monotonicline',
    'octagramspiral': 'octagramspiral',
    'rectilinear': 'zig-zag',
    'rectilineargapfill': 'zig-zag',
    'rectiwithperimeter': 'zig-zag',
    'sawtooth': 'zig-zag',
    'scatteredrectilinear': 'zig-zag',
    'smooth': 'monotonic',
    'smoothhilbert': 'hilbertcurve',
    'smoothtriple': 'triangles',
    'stars': 'tri-hexagon',
    'supportcubic': 'supportcubic',
    'triangles': 'triangles',
}

# Support style mappings
SUPPORT_STYLES = {
    'grid': ('normal', 'grid'),
    'snug': ('normal', 'snug'),
    'tree': ('tree', 'default'),
    'organic': ('tree', 'organic'),
}

# Support pattern mappings
SUPPORT_PATTERNS = {
    'rectilinear', 'rectilinear-grid', 'honeycomb', 'lightning', 'default', 'hollow'
}

# Interface pattern mappings
INTERFACE_PATTERNS = {
    'auto', 'rectilinear', 'concentric', 'rectilinear_interlaced', 'grid'
}

# GCode flavor mappings
GCODE_FLAVORS = {
    'klipper': 'klipper',
    'mach3': 'reprapfirmware',
    'machinekit': 'reprapfirmware',
    'makerware': 'reprapfirmware',
    'marlin': 'marlin',
    'marlin2': 'marlin2',
    'no-extrusion': 'reprapfirmware',
    'repetier': 'reprapfirmware',
    'reprap': 'reprapfirmware',
    'reprapfirmware': 'reprapfirmware',
    'sailfish': 'reprapfirmware',
    'smoothie': 'reprapfirmware',
    'teacup': 'reprapfirmware',
    'sprinter': 'reprapfirmware',
}

# Host type mappings
HOST_TYPES = {
    'repetier': 'repetier',
    'prusalink': 'prusalink',
    'prusaconnect': 'prusaconnect',
    'octoprint': 'octoprint',
    'moonraker': 'octoprint',
    'mks': 'mks',
    'klipper': 'octoprint',
    'flashair': 'flashair',
    'duet': 'duet',
    'astrobox': 'astrobox',
}

# Z-hop enforcement mappings
ZHOP_ENFORCEMENT = {
    'All surfaces': 'All Surfaces',
    'Not on top': 'Bottom Only',
    'Only on top': 'Top Only',
}

# Thumbnail format mappings
THUMBNAIL_FORMAT = {
    'PNG': 'PNG',
    'JPG': 'JPG',
    'QOI': 'QOI',
    'BIQU': 'BTT_TFT',
}


class SlicerConverter:
    """Converts PrusaSlicer/SuperSlicer profiles to OrcaSlicer format."""

    def __init__(self):
        self.status = {
            'force_out': False,
            'max_temp': 0,
            'interactive_mode': False,
            'slicer_flavor': None,
            'ini_type': None,
            'profile_name': None,
            'ironing_type': None,
            'dirs': {
                'output': None,
                'data': self._get_data_dir(),
                'slicer': None,
                'temp': None,
            },
            'to_var': {
                'external_perimeters_first': None,
                'infill_first': None,
                'ironing': None,
            },
            'value': {
                'on_existing': None,
                'physical_printer': None,
                'nozzle_size': None,
                'inherits': None,
            },
        }
        self.new_hash: Dict[str, Any] = {}
        self.converted_files: Dict[str, List] = {}
        self.source_hash: Dict[str, Any] = {}

    def _get_data_dir(self) -> Path:
        """Get the default data directory based on OS."""
        system = platform.system()
        home = Path.home()

        dirs = SYSTEM_DIRECTORIES['os'].get(system, ['.config'])
        return home / Path(*dirs)

    @staticmethod
    def is_decimal(value: Any) -> bool:
        """Check if a value is a decimal number."""
        if value is None:
            return False
        try:
            s = str(value).strip()
            if s.endswith('%'):
                return False
            float(s)
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def is_percent(value: Any) -> bool:
        """Check if a value is a percentage."""
        if value is None:
            return False
        s = str(value).strip()
        if not s.endswith('%'):
            return False
        try:
            float(s[:-1])
            return True
        except (ValueError, TypeError):
            return False

    @staticmethod
    def remove_percent(value: Any) -> str:
        """Remove the percent symbol from a value."""
        if value is None:
            return ''
        return str(value).strip().rstrip('%')

    @staticmethod
    def multivalue_to_array(input_string: Optional[str]) -> List[str]:
        """Convert comma or semicolon-separated string to array."""
        if not input_string:
            return []
        delimiter = ',' if ',' in input_string else ';'
        return [x.strip() for x in input_string.split(delimiter)]

    def is_config_bundle(self, file_path: Path) -> bool:
        """Check if input file is a config bundle."""
        content = file_path.read_text(encoding='utf-8', errors='replace')
        bundle_detected = bool(re.search(r'\[\w+:[\w\s\+\-]+\]', content))
        if bundle_detected:
            self.status['dirs']['temp'] = Path(tempfile.mkdtemp())
            self.status['dirs']['slicer'] = self.status['dirs']['temp']
            (self.status['dirs']['slicer'] / 'physical_printer').mkdir(exist_ok=True)
        return bundle_detected

    def process_config_bundle(self, file_path: Path) -> List[Path]:
        """Process a config bundle and split into individual INI files."""
        content = file_path.read_text(encoding='utf-8', errors='replace')
        header_match = re.search(r'^(# generated[^\n]*)', content, re.MULTILINE)
        header_line = header_match.group(1) if header_match else ''

        file_objects = []
        pattern = r'\[([\w\s\+\-]+):([^\]]+)\]\n(.*?)(?=\n\[|$)'

        for match in re.finditer(pattern, content, re.DOTALL):
            profile_type, profile_name, profile_content = match.groups()
            physical_printer_profile = profile_type == 'physical_printer'

            # Clean filename of illegal characters
            temp_filename = re.sub(r'[<>:"/\\|?*\x00-\x1F:]', '', profile_name)

            temp_dir = self.status['dirs']['temp']
            if (temp_dir / f'{temp_filename}.ini').exists():
                temp_filename = f'{profile_type}_{temp_filename}'

            if physical_printer_profile:
                temp_file = self.status['dirs']['slicer'] / 'physical_printer' / f'{temp_filename}.ini'
            else:
                temp_file = temp_dir / f'{temp_filename}.ini'

            temp_file.write_text(
                f"{header_line}\n\n"
                f"ini_type = {profile_type}\n"
                f"profile_name = {profile_name}\n"
                f"{profile_content}",
                encoding='utf-8'
            )

            if not physical_printer_profile:
                file_objects.append(temp_file)

        return file_objects

    def evaluate_print_order(self, external_perimeters_first: bool, infill_first: bool) -> str:
        """Translate the feature print sequence settings."""
        if not external_perimeters_first and not infill_first:
            return "inner wall/outer wall/infill"
        if external_perimeters_first and not infill_first:
            return "outer wall/inner wall/infill"
        if not external_perimeters_first and infill_first:
            return "infill/inner wall/outer wall"
        if external_perimeters_first and infill_first:
            return "infill/outer wall/inner wall"
        return "inner wall/outer wall/infill"

    def evaluate_ironing_type(self, ironing: bool, ironing_type: Optional[str]) -> str:
        """Translate the ironing type settings."""
        if ironing:
            return ironing_type if ironing_type else "no ironing"
        return "no ironing"

    def percent_to_float(self, value: Any) -> str:
        """Convert percentage to float."""
        if not self.is_percent(value):
            return str(value) if value is not None else ''
        new_float = float(self.remove_percent(value)) / 100
        return '2' if new_float > 2 else str(new_float)

    def percent_to_mm(self, mm_comparator: Any, percent_param: Any) -> Optional[str]:
        """Convert percentage value to millimeters."""
        if mm_comparator is None or percent_param is None:
            return None
        mm_str = str(mm_comparator).strip()
        percent_str = str(percent_param).strip()
        if mm_str == '' or percent_str == '':
            return None
        if not self.is_percent(percent_str):
            return percent_str
        if self.is_percent(mm_str):
            return None
        try:
            result = float(mm_str) * (float(self.remove_percent(percent_str)) / 100)
            return str(result)
        except (ValueError, TypeError):
            return None

    def mm_to_percent(self, mm_comparator: Any, mm_param: Any) -> Optional[str]:
        """Convert millimeter values to percentage."""
        if self.is_percent(mm_param):
            return str(mm_param)
        if self.is_percent(mm_comparator):
            return None
        try:
            return f"{(float(mm_param) / float(mm_comparator)) * 100}%"
        except (ValueError, ZeroDivisionError, TypeError):
            return None

    def unbackslash_gcode(self, value: Any) -> List[str]:
        """Process gcode by removing quotes and handling escape sequences."""
        if value is None:
            return ['']
        s = str(value)
        # Remove surrounding quotes
        if s.startswith('"') and s.endswith('"'):
            s = s[1:-1]
        # Handle escape sequences
        s = s.replace('\\n', '\n')
        s = s.replace('\\t', '\t')
        s = s.replace('\\\\', '\\')
        return [s]

    def ini_reader(self, file_path: Path) -> Dict[str, Any]:
        """Parse an INI file into a dictionary."""
        config = {}
        content = file_path.read_text(encoding='utf-8', errors='replace')

        for line in content.split('\n'):
            # Detect slicer flavor
            match = re.match(r'^#\s*generated\s+by\s+(\S+)', line, re.IGNORECASE)
            if match:
                self.status['slicer_flavor'] = match.group(1)
                continue

            # Skip empty and comment lines
            if re.match(r'^\s*(?:#|$)', line):
                continue

            # Parse key = value
            parts = line.split('=', 1)
            if len(parts) == 2:
                key = parts[0].strip()
                value = parts[1].strip()
                config[key] = value

        if self.status['slicer_flavor']:
            self.status['dirs']['slicer'] = self.status['dirs']['data'] / self.status['slicer_flavor']

        return config

    def detect_ini_type(self, source_ini: Dict[str, Any]) -> Optional[str]:
        """Detect the INI file type from its contents."""
        if 'ini_type' in source_ini:
            return source_ini['ini_type'].lower()

        # Count matching parameters for each type
        type_counts = {}
        for ini_type, params in PARAMETER_MAP.items():
            if ini_type == 'physical_printer':
                continue
            type_counts[ini_type] = sum(1 for p in source_ini if p in params)

        # Return None if all counts are less than 10
        if all(count < 10 for count in type_counts.values()):
            return None

        # Return type with highest count
        return max(type_counts, key=type_counts.get)

    def convert_params(self, parameter: str, source_ini: Dict[str, Any]) -> Any:
        """Convert a parameter value to OrcaSlicer format."""
        new_value = source_ini.get(parameter)
        if new_value is None or new_value == 'nil':
            return None

        ini_type = self.status['ini_type']
        slicer_flavor = self.status['slicer_flavor'] or 'PrusaSlicer'

        # Handle multivalue parameters
        if parameter in MULTIVALUE_PARAMS:
            arr = self.multivalue_to_array(new_value)
            if MULTIVALUE_PARAMS[parameter] == 'single':
                new_value = arr[0] if arr else new_value
            else:
                new_value = arr

        # Get default_speed for SuperSlicer
        default_speed = source_ini.get('default_speed') if slicer_flavor == 'SuperSlicer' else None

        # Handle gcode blocks
        gcode_params = [
            'start_filament_gcode', 'end_filament_gcode', 'post_process',
            'before_layer_gcode', 'toolchange_gcode', 'layer_gcode',
            'feature_gcode', 'end_gcode', 'pause_print_gcode', 'start_gcode',
            'template_custom_gcode', 'notes', 'filament_notes', 'printer_notes'
        ]
        if parameter in gcode_params:
            return self.unbackslash_gcode(new_value)

        # Filament type conversion
        if parameter == 'filament_type':
            return FILAMENT_TYPES.get(new_value, new_value)

        # Max volumetric speed
        if parameter == 'filament_max_volumetric_speed':
            try:
                if float(new_value) > 0:
                    return str(new_value)
            except (ValueError, TypeError):
                pass
            filament_type = source_ini.get('filament_type', 'PLA')
            return DEFAULT_MVS.get(filament_type, '15')

        # Draft shield
        if parameter == 'draft_shield':
            if new_value == 'disabled':
                return '0'
            elif new_value == 'enabled':
                return '1'
            return new_value

        # External perimeter fan speed
        if parameter == 'external_perimeter_fan_speed':
            try:
                val = int(new_value)
                return '0%' if val < 0 else f'{val}%'
            except (ValueError, TypeError):
                return '0%'

        # Ironing type tracking
        if parameter == 'ironing_type':
            self.status['ironing_type'] = new_value
            return new_value

        # Default filament profile
        if parameter == 'default_filament_profile':
            arr = self.multivalue_to_array(new_value)
            if arr:
                result = self.unbackslash_gcode(arr[0])
                return result[0] if result else ''
            return ''

        # Retract lift top
        if parameter == 'retract_lift_top':
            arr = self.multivalue_to_array(new_value)
            if arr:
                val = self.unbackslash_gcode(arr[0])
                return ZHOP_ENFORCEMENT.get(val[0] if val else '', val[0] if val else '')
            return ''

        # Percent to mm conversions
        percent_to_mm_params = {
            'max_layer_height': 'nozzle_size',
            'min_layer_height': 'nozzle_size',
            'fuzzy_skin_point_dist': 'nozzle_size',
            'fuzzy_skin_thickness': 'nozzle_size',
            'small_perimeter_min_length': 'nozzle_size',
        }
        if parameter in percent_to_mm_params:
            comparator = self.status['value'].get(percent_to_mm_params[parameter])
            return self.percent_to_mm(comparator, new_value)

        # Percent to float conversions
        if parameter in ['bridge_flow_ratio', 'fill_top_flow_ratio', 'first_layer_flow_ratio']:
            return self.percent_to_float(new_value)

        # Wall transition length
        if parameter == 'wall_transition_length':
            nozzle = self.status['value'].get('nozzle_size')
            return self.mm_to_percent(nozzle, new_value)

        # Machine limits usage
        if parameter == 'machine_limits_usage':
            return '1' if new_value == 'emit_to_gcode' else '0'

        # Remaining times (inverted for disable_m73)
        if parameter == 'remaining_times':
            return '0' if new_value == '1' else '1'

        # Support bottom contact distance
        if parameter == 'support_material_bottom_contact_distance':
            if new_value == '0':
                return source_ini.get('support_material_contact_distance', '0')
            return new_value

        # Support style
        if parameter == 'support_material_style':
            if new_value in SUPPORT_STYLES:
                support_type, support_style = SUPPORT_STYLES[new_value]
                genstyle = 'auto' if source_ini.get('support_material_auto', '0') == '1' else 'manual'
                self.new_hash['support_type'] = f'{support_type}({genstyle})'
                self.new_hash['support_style'] = support_style
            return None

        # Infill types
        if parameter in ['fill_pattern', 'top_fill_pattern', 'bottom_fill_pattern', 'solid_fill_pattern']:
            return INFILL_TYPES.get(new_value, new_value)

        # GCode flavor
        if parameter == 'gcode_flavor':
            return GCODE_FLAVORS.get(new_value)

        # Host type
        if parameter == 'host_type':
            return HOST_TYPES.get(new_value)

        # Thumbnails format
        if parameter == 'thumbnails_format':
            return THUMBNAIL_FORMAT.get(new_value, new_value)

        # Support pattern
        if parameter == 'support_material_pattern':
            return new_value if new_value in SUPPORT_PATTERNS else 'default'

        # Support interface pattern
        if parameter == 'support_material_interface_pattern':
            return new_value if new_value in INTERFACE_PATTERNS else 'auto'

        # Seam position
        if parameter == 'seam_position':
            return SEAM_POSITIONS.get(new_value, new_value)

        # Independent support layer height
        if parameter == 'support_material_layer_height':
            try:
                return '1' if float(new_value) > 0 else '0'
            except (ValueError, TypeError):
                return '0'

        # Output filename format
        if parameter == 'output_filename_format':
            return re.sub(r'[\[\]]', lambda m: '{' if m.group() == '[' else '}', new_value)

        # Support XY spacing
        if parameter == 'support_material_xy_spacing':
            ext_width = source_ini.get('external_perimeter_extrusion_width')
            result = self.percent_to_mm(ext_width, new_value)
            if result is None:
                result = self.percent_to_mm(self.status['value'].get('nozzle_size'), new_value)
            return result

        # Empty extrusion width
        if parameter == 'extrusion_width':
            return '0' if new_value == '' else new_value

        # Infill every layers
        if parameter == 'infill_every_layers':
            try:
                return '1' if int(new_value) > 0 else '0'
            except (ValueError, TypeError):
                return '0'

        # Complete objects
        if parameter == 'complete_objects':
            return 'by object' if new_value in ('1', 'true') else 'by layer'

        # Speed conversions
        speed_conversions = {
            'external_perimeter_speed': lambda: self.percent_to_mm(
                source_ini.get('perimeter_speed'), new_value),
            'first_layer_speed': lambda: self.percent_to_mm(
                source_ini.get('perimeter_speed'), new_value),
            'top_solid_infill_speed': lambda: self.percent_to_mm(
                self.new_hash.get('internal_solid_infill_speed'), new_value),
            'support_material_interface_speed': lambda: self.percent_to_mm(
                source_ini.get('support_material_speed'), new_value),
        }

        if parameter in speed_conversions:
            result = speed_conversions[parameter]()
            return result if result else new_value

        if parameter == 'first_layer_infill_speed':
            if slicer_flavor == 'PrusaSlicer':
                return self.percent_to_mm(source_ini.get('infill_speed'),
                                          source_ini.get('first_layer_speed'))
            return self.percent_to_mm(source_ini.get('infill_speed'), new_value)

        if parameter == 'solid_infill_speed':
            ref = source_ini.get('infill_speed') if slicer_flavor == 'PrusaSlicer' else default_speed
            return self.percent_to_mm(ref, new_value)

        superslicer_speed_params = ['perimeter_speed', 'support_material_speed', 'bridge_speed']
        if parameter in superslicer_speed_params and slicer_flavor == 'SuperSlicer':
            return self.percent_to_mm(default_speed, new_value)

        if parameter == 'infill_speed' and slicer_flavor == 'SuperSlicer':
            return self.percent_to_mm(self.new_hash.get('internal_solid_infill_speed'), new_value)

        if parameter == 'small_perimeter_speed' and slicer_flavor == 'SuperSlicer':
            return self.percent_to_mm(self.new_hash.get('sparse_infill_speed'), new_value)

        if parameter == 'gap_fill_speed' and slicer_flavor == 'SuperSlicer':
            return self.percent_to_mm(self.new_hash.get('sparse_infill_speed'), new_value)

        return new_value

    def calculate_print_params(self, source_ini: Dict[str, Any]) -> None:
        """Calculate and set print-specific parameters."""
        # Process speed sequence
        for parameter in SPEED_SEQUENCE:
            if parameter not in source_ini:
                continue
            new_value = self.convert_params(parameter, source_ini)
            if new_value is None:
                continue
            # Limit to one decimal place
            if self.is_decimal(new_value):
                try:
                    new_value = f'{float(new_value):.1f}'.rstrip('0').rstrip('.')
                except (ValueError, TypeError):
                    pass
            self.new_hash[SPEED_PARAMS[parameter]] = str(new_value)

        # Dynamic overhang speeds
        enable_dynamic = source_ini.get('enable_dynamic_overhang_speeds', '0') == '1'
        self.new_hash['enable_overhang_speed'] = '1' if enable_dynamic else '0'

        if enable_dynamic:
            if self.status['slicer_flavor'] == 'SuperSlicer':
                speeds = self.multivalue_to_array(source_ini.get('dynamic_overhang_speeds', ''))
            else:
                speeds = [
                    source_ini.get('overhang_speed_0', ''),
                    source_ini.get('overhang_speed_1', ''),
                    source_ini.get('overhang_speed_2', ''),
                    source_ini.get('overhang_speed_3', ''),
                ]
            overhang_keys = ['overhang_1_4_speed', 'overhang_2_4_speed',
                            'overhang_3_4_speed', 'overhang_4_4_speed']
            if len(speeds) >= 4:
                for i, key in enumerate(overhang_keys):
                    self.new_hash[key] = speeds[3 - i]

        # Wall infill order
        ext_first = source_ini.get('external_perimeters_first', '0') in ('1', 'true')
        inf_first = source_ini.get('infill_first', '0') in ('1', 'true')
        self.new_hash['wall_infill_order'] = self.evaluate_print_order(ext_first, inf_first)

        # Ironing type
        ironing = source_ini.get('ironing', '0') in ('1', 'true')
        self.new_hash['ironing_type'] = self.evaluate_ironing_type(ironing, self.status['ironing_type'])

    def convert_profile(self, file_path: Path, nozzle_size: Optional[str] = None,
                        on_existing: Optional[str] = None) -> Tuple[Optional[Dict], str, str]:
        """Convert a single profile file to OrcaSlicer format."""
        # Parse the INI file
        source_ini = self.ini_reader(file_path)

        if not self.status['slicer_flavor']:
            return None, 'unsupported', file_path.stem

        # Detect INI type
        ini_type = self.status['ini_type'] or self.detect_ini_type(source_ini)
        if not ini_type:
            return None, 'unsupported', file_path.stem

        self.status['ini_type'] = ini_type
        profile_name = source_ini.get('profile_name', file_path.stem)
        self.status['profile_name'] = profile_name

        # Set nozzle size
        if 'nozzle_diameter' in source_ini:
            nozzle_diameters = self.multivalue_to_array(source_ini['nozzle_diameter'])
            self.status['value']['nozzle_size'] = nozzle_diameters[0] if nozzle_diameters else nozzle_size
        elif nozzle_size:
            self.status['value']['nozzle_size'] = nozzle_size
        elif ini_type == 'print' and 'layer_height' in source_ini:
            try:
                self.status['value']['nozzle_size'] = str(2 * float(source_ini['layer_height']))
            except (ValueError, TypeError):
                self.status['value']['nozzle_size'] = '0.4'

        # Initialize new hash
        self.new_hash = {}

        # Get parameter map for this type
        param_map = PARAMETER_MAP.get(ini_type, {})

        # Track combination settings
        for param in ['external_perimeters_first', 'infill_first', 'ironing']:
            if param in source_ini:
                self.status['to_var'][param] = source_ini[param] in ('1', 'true')

        # Process each parameter
        for parameter in source_ini:
            if parameter == 'profile_name':
                continue
            if parameter not in param_map:
                continue

            new_value = self.convert_params(parameter, source_ini)
            if new_value is None:
                continue

            target_key = param_map[parameter]

            # Handle multi-target mappings
            if isinstance(target_key, list):
                for tk in target_key:
                    self.new_hash[tk] = new_value if not isinstance(new_value, list) else new_value
            else:
                self.new_hash[target_key] = new_value

            # Track max temperature
            if parameter in ('first_layer_temperature', 'temperature'):
                try:
                    temp = float(new_value) if not isinstance(new_value, list) else float(new_value[0])
                    if temp > self.status['max_temp']:
                        self.status['max_temp'] = temp
                except (ValueError, TypeError):
                    pass

        # Add metadata
        self.new_hash.update({
            f'{ini_type}_settings_id': profile_name,
            'name': profile_name,
            'from': 'User',
            'is_custom_defined': '1',
            'version': ORCA_SLICER_VERSION,
        })

        # Profile-specific processing
        if ini_type == 'filament':
            self.new_hash['nozzle_temperature_range_low'] = '0'
            self.new_hash['nozzle_temperature_range_high'] = str(int(self.status['max_temp']))
            if 'slowdown_below_layer_time' in source_ini:
                try:
                    slow_down = float(source_ini['slowdown_below_layer_time']) > 0
                    self.new_hash['slow_down_for_layer_cooling'] = '1' if slow_down else '0'
                except (ValueError, TypeError):
                    pass

        elif ini_type == 'print':
            self.calculate_print_params(source_ini)

        return self.new_hash, ini_type, profile_name

    def write_json(self, output_dir: Path, data: Dict, filename: str) -> Path:
        """Write the converted data to a JSON file."""
        output_path = output_dir / f"{filename}.json"

        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False, sort_keys=True)

        return output_path


def select_files_interactive(data_dir: Path) -> Tuple[List[Path], Optional[str]]:
    """Interactive file selection."""
    # Find slicer directories
    slicers = []
    for name in ['PrusaSlicer', 'SuperSlicer']:
        slicer_dir = data_dir / name
        if slicer_dir.exists():
            slicers.append(name)

    if not slicers:
        print(f"No PrusaSlicer or SuperSlicer directories found in {data_dir}")
        return [], None

    print("\nAvailable slicers:")
    for i, s in enumerate(slicers, 1):
        print(f"  {i}. {s}")

    try:
        choice = input("\nSelect slicer (number): ").strip()
        slicer = slicers[int(choice) - 1]
    except (ValueError, IndexError):
        print("Invalid selection")
        return [], None

    slicer_dir = data_dir / slicer

    # Select profile type
    profile_types = []
    for ptype in ['filament', 'print', 'printer']:
        if (slicer_dir / ptype).exists():
            profile_types.append(ptype.capitalize())

    print(f"\nProfile types in {slicer}:")
    for i, pt in enumerate(profile_types, 1):
        print(f"  {i}. {pt}")

    try:
        choice = input("\nSelect profile type (number): ").strip()
        profile_type = profile_types[int(choice) - 1].lower()
    except (ValueError, IndexError):
        print("Invalid selection")
        return [], None

    # Select profiles
    profile_dir = slicer_dir / profile_type
    ini_files = sorted(profile_dir.glob('*.ini'))

    if not ini_files:
        print(f"No INI files found in {profile_dir}")
        return [], None

    print(f"\nAvailable {profile_type} profiles:")
    for i, f in enumerate(ini_files, 1):
        print(f"  {i}. {f.stem}")
    print(f"  A. All profiles")

    selection = input("\nSelect profiles (comma-separated numbers or 'A' for all): ").strip()

    if selection.upper() == 'A':
        return ini_files, profile_type

    try:
        indices = [int(x.strip()) - 1 for x in selection.split(',')]
        return [ini_files[i] for i in indices if 0 <= i < len(ini_files)], profile_type
    except (ValueError, IndexError):
        print("Invalid selection")
        return [], None


def main():
    parser = argparse.ArgumentParser(
        description='Convert PrusaSlicer/SuperSlicer profiles to OrcaSlicer format.',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  %(prog)s --input profile.ini
  %(prog)s --input "*.ini" --outdir ~/OrcaSlicer
  %(prog)s --input config_bundle.ini --nozzle-size 0.4
        """
    )

    parser.add_argument(
        '--input', '-i',
        nargs='+',
        help='Input INI file(s). Supports wildcards.'
    )
    parser.add_argument(
        '--outdir', '-o',
        help='Output directory for OrcaSlicer profiles.'
    )
    parser.add_argument(
        '--nozzle-size',
        help='Nozzle diameter in mm (for print profiles).'
    )
    parser.add_argument(
        '--on-existing',
        choices=['skip', 'overwrite', 'merge'],
        default='overwrite',
        help='Behavior when output file exists (default: overwrite).'
    )
    parser.add_argument(
        '--force-output',
        action='store_true',
        help='Force output to specified directory without subdirectories.'
    )
    parser.add_argument(
        '--physical-printer',
        help='Physical printer INI file for printer profiles.'
    )

    args = parser.parse_args()

    converter = SlicerConverter()

    # Set output directory
    if args.outdir:
        converter.status['dirs']['output'] = Path(args.outdir)
    else:
        converter.status['dirs']['output'] = converter.status['dirs']['data'] / 'OrcaSlicer'

    converter.status['force_out'] = args.force_output

    # Get input files
    input_files = []
    ini_type = None

    if args.input:
        for pattern in args.input:
            # Expand wildcards
            expanded = glob(pattern)
            if expanded:
                input_files.extend(Path(f) for f in expanded if Path(f).suffix == '.ini')
            else:
                p = Path(pattern)
                if p.exists() and p.suffix == '.ini':
                    input_files.append(p)
                elif p.is_dir():
                    input_files.extend(p.glob('*.ini'))
                else:
                    print(f"Warning: File not found or not an INI file: {pattern}")
    else:
        # Interactive mode
        converter.status['interactive_mode'] = True
        input_files, ini_type = select_files_interactive(converter.status['dirs']['data'])
        if ini_type:
            converter.status['ini_type'] = ini_type

    if not input_files:
        print("No input files specified or found.")
        sys.exit(1)

    # Expand config bundles
    expanded_files = []
    for input_file in input_files:
        if converter.is_config_bundle(input_file):
            print(f"Detected config bundle: {input_file}")
            expanded_files.extend(converter.process_config_bundle(input_file))
        else:
            expanded_files.append(input_file)

    # Process each file
    results = []
    for input_file in expanded_files:
        print(f"\nProcessing: {input_file.name}")

        try:
            # Get nozzle size for print profiles
            nozzle_size = args.nozzle_size
            if not nozzle_size and converter.status['interactive_mode']:
                temp_ini = converter.ini_reader(input_file)
                detected_type = converter.detect_ini_type(temp_ini)
                if detected_type == 'print' and 'nozzle_diameter' not in temp_ini:
                    nozzle_size = input(f"  Enter nozzle size (mm) for {input_file.stem} [0.4]: ").strip()
                    if not nozzle_size:
                        nozzle_size = '0.4'
                # Reset for next file
                converter.status['slicer_flavor'] = None

            data, detected_ini_type, profile_name = converter.convert_profile(
                input_file, nozzle_size, args.on_existing)

            if data is None:
                print(f"  Skipped: unsupported file type")
                continue

            # Determine output path
            output_dir = converter.status['dirs']['output']
            if not converter.status['force_out']:
                subdir = SYSTEM_DIRECTORIES['output'].get(detected_ini_type, ['user', 'default', 'filament'])
                output_dir = output_dir / Path(*subdir)

            output_dir.mkdir(parents=True, exist_ok=True)

            # Handle existing files
            output_file = output_dir / f"{profile_name}.json"
            status = 'converted'

            if output_file.exists():
                if args.on_existing == 'skip':
                    print(f"  Skipped (exists): {output_file}")
                    continue
                elif args.on_existing == 'merge':
                    with open(output_file, 'r', encoding='utf-8') as f:
                        existing = json.load(f)
                    for k, v in data.items():
                        if k not in existing:
                            existing[k] = v
                    data = existing
                    status = 'merged'

            # Write output
            written = converter.write_json(output_dir, data, profile_name)
            results.append({
                'input': str(input_file),
                'output': str(written),
                'type': detected_ini_type,
                'name': profile_name,
                'status': status,
            })
            print(f"  {status.capitalize()}: {profile_name} ({detected_ini_type}) -> {written.name}")

            # Reset for next file
            converter.status['ini_type'] = ini_type  # Keep selected type in interactive mode
            converter.status['max_temp'] = 0
            converter.status['ironing_type'] = None

        except Exception as e:
            print(f"  Error: {e}")
            import traceback
            traceback.print_exc()

    # Summary
    print(f"\n{'=' * 60}")
    print(f"Conversion complete: {len(results)} profile(s) processed")
    for r in results:
        print(f"  [{r['type']:8}] {r['name']} ({r['status']})")


if __name__ == '__main__':
    main()
