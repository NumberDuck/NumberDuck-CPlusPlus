cmake_minimum_required(VERSION 3.14...3.22)

project(
	NumberDuck
	VERSION 3.0.7
	LANGUAGES CXX
)

if(PROJECT_SOURCE_DIR STREQUAL PROJECT_BINARY_DIR)
	message(
		FATAL_ERROR
			"In-source builds not allowed. Please make a new directory (called a build directory) and run CMake from there."
	)
endif()


add_library(${PROJECT_NAME} NumberDuck.cpp)


if(SIMPLE_OUTPUT)
	include(CMakePackageConfigHelpers)
	include(GNUInstallDirs)

	install(
		FILES "${PROJECT_SOURCE_DIR}/NumberDuck.hpp"
		DESTINATION "${CMAKE_INSTALL_INCLUDEDIR}"
	)
	install(
		TARGETS ${PROJECT_NAME}
		LIBRARY DESTINATION ${CMAKE_INSTALL_LIBDIR}
		ARCHIVE DESTINATION ${CMAKE_INSTALL_LIBDIR}
	)
else()
	include(cmake/CPM.cmake)
	CPMAddPackage("gh:TheLartians/PackageProject.cmake@1.13.0")
	string(TOLOWER ${PROJECT_NAME}/version.h VERSION_HEADER_LOCATION)

	packageProject(
		NAME ${PROJECT_NAME}
		VERSION ${PROJECT_VERSION}
		NAMESPACE ${PROJECT_NAME}
		BINARY_DIR ${PROJECT_BINARY_DIR}
		INCLUDE_DESTINATION include/${PROJECT_NAME}-${PROJECT_VERSION}
		VERSION_HEADER "${VERSION_HEADER_LOCATION}"
		COMPATIBILITY SameMajorVersion
	)

	install(
		FILES "${PROJECT_SOURCE_DIR}/NumberDuck.hpp"
		DESTINATION "include/${PROJECT_NAME}-${PROJECT_VERSION}"
		COMPONENT "${PROJECT_NAME}_Development"
	)
endif()