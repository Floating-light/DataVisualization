#pragma once
#include <qheaderview.h>
#include <QPainter>
#include <QMouseEvent>
#include <QCheckbox>

class CheckHeaderView :
	public QHeaderView
{
	Q_OBJECT
public:
	CheckHeaderView(Qt::Orientation orientation, QWidget* parent);
	~CheckHeaderView();

signals:

	void headCheckBoxToggled(bool checked);

protected:

	void paintSection(QPainter* painter, const QRect& rect, int logicalIndex) const;

	void mousePressEvent(QMouseEvent* event);

	void mouseMoveEvent(QMouseEvent* event);

private:

	bool m_isOn;//是否选中

	QPoint m_mousePoint;//鼠标位置

	mutable QRect m_RectHeaderCheckBox;//复选框的位置
};

